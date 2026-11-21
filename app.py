import streamlit as st
import pandas as pd
import xmlrpc.client
import base64
import os
import io
from datetime import datetime
from docx import Document as DocxDocument
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn

import odoo_rules as rules

# ─────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────
ODOO_URL     = os.environ.get("ODOO_URL",     "https://inovues.odoo.com")
ODOO_DB      = os.environ.get("ODOO_DB",      "inovues")
ODOO_USER    = os.environ.get("ODOO_USER",    "sketterer@inovues.com")
ODOO_API_KEY = os.environ.get("ODOO_API_KEY", "")

LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ilogo.png")  # INOVUES logo for PO header
LOGO_WIDTH_IN = 1.9  # logo width on the PO, in inches (horizontal logo)

# INOVUES brand teal, sampled from ilogo.png. Used for both PO table header rows (white text on top).
BRAND_TEAL = "118DA1"

SHIP_TO_DEFAULT = "Momentum Glass, LLC\nAttn: INOVUES, INC.\n25825 Aldine Westfield Rd.\nSpring, TX 77373\n281.809.2830"

sq_inches_to_sq_feet = 1 / 144

st.set_page_config(
    page_title="INOVUES PO Generator",
    page_icon="📝",
    layout="wide"
)

# ─────────────────────────────────────────────────────────────
# ODOO HELPERS
# ─────────────────────────────────────────────────────────────
def get_odoo_connection():
    common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common")
    uid = common.authenticate(ODOO_DB, ODOO_USER, ODOO_API_KEY, {})
    if not uid:
        raise Exception("Odoo authentication failed — check ODOO_API_KEY.")
    models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")
    return uid, models

def odoo_call(models, uid, model, method, args, kwargs={}):
    return models.execute_kw(ODOO_DB, uid, ODOO_API_KEY, model, method, args, kwargs)

@st.cache_data(ttl=120, show_spinner="Loading projects from Odoo...")
def fetch_projects():
    try:
        uid, models = get_odoo_connection()
        projects = odoo_call(models, uid, "project.project", "search_read",
            [[("active", "=", True)]],
            {"fields": ["id", "name"], "order": "name asc", "limit": 200})
        return {p["name"]: p["id"] for p in projects}, None
    except Exception as e:
        return {}, str(e)

@st.cache_data(ttl=60, show_spinner="Fetching files from Odoo...")
def fetch_project_attachments(project_id, po_type_key="Glass"):
    """Get attachments from the relevant tasks in this project.

    For Glass POs, looks for SWR Cutlist tasks (which produce *_SWR_Glass_*.xlsx).
    For Aluminium POs, looks for Cutting Optimization tasks (which produce cutting_list_*.xlsx).
    """
    try:
        uid, models = get_odoo_connection()

        # Choose task name keyword based on PO type
        task_keyword = "Cutting Optimization" if po_type_key == "Aluminium" else "SWR"

        # Find tasks in this project matching the keyword
        tasks = odoo_call(models, uid, "project.task", "search_read",
            [[("project_id", "=", project_id), ("name", "ilike", task_keyword)]],
            {"fields": ["id", "name", "create_date"], "order": "create_date desc", "limit": 10})

        if not tasks:
            # Fall back to any task in the project
            tasks = odoo_call(models, uid, "project.task", "search_read",
                [[("project_id", "=", project_id)]],
                {"fields": ["id", "name", "create_date"], "order": "create_date desc", "limit": 20})

        if not tasks:
            return [], None

        # Get attachments from all found tasks
        task_ids = [t["id"] for t in tasks]
        attachments = odoo_call(models, uid, "ir.attachment", "search_read",
            [[("res_model", "=", "project.task"), ("res_id", "in", task_ids),
               ("name", "ilike", ".xlsx")]],
            {"fields": ["id", "name", "res_id", "datas", "create_date"],
             "order": "create_date desc"})

        # Add task name context
        task_map = {t["id"]: t["name"] for t in tasks}
        for a in attachments:
            a["task_name"] = task_map.get(a["res_id"], "Unknown task")

        return attachments, None
    except Exception as e:
        return [], str(e)

@st.cache_data(ttl=30, show_spinner="Loading vendors from Odoo...")
def fetch_vendors():
    try:
        uid, models = get_odoo_connection()
        vendors = odoo_call(models, uid, "res.partner", "search_read",
            [[("supplier_rank", ">", 0)]],
            {"fields": ["id", "name", "email", "phone", "street", "street2",
                        "city", "state_id", "zip", "country_id", "child_ids"],
             "order": "name asc", "limit": 200})
        if not vendors:
            vendors = odoo_call(models, uid, "res.partner", "search_read",
                [[("is_company", "=", True)]],
                {"fields": ["id", "name", "email", "phone", "street", "street2",
                            "city", "state_id", "zip", "country_id", "child_ids"],
                 "order": "name asc", "limit": 200})
        # Enrich with address + contact
        for v in vendors:
            addr_parts = [p for p in [v.get("street"), v.get("street2")] if p]
            city_line = ", ".join(p for p in [
                v.get("city"),
                v.get("state_id", [False, ""])[1] if isinstance(v.get("state_id"), list) else "",
                v.get("zip")
            ] if p)
            if city_line:
                addr_parts.append(city_line)
            country = v.get("country_id", [False, ""])[1] if isinstance(v.get("country_id"), list) else ""
            if country:
                addr_parts.append(country)
            v["full_address"] = "\n".join(addr_parts)

            # Pick contact: prefer child tagged 'Orders', else first child.
            # The 'Orders' tag (res.partner.category) flags the primary ordering
            # contact for a vendor — set it on whichever child should appear in POs.
            v["contact_name"] = ""
            if v.get("child_ids"):
                primary_ids = odoo_call(models, uid, "res.partner", "search",
                    [[("id", "in", v["child_ids"]),
                      ("category_id.name", "=", "Orders")]],
                    {"limit": 1})
                target_ids = primary_ids if primary_ids else v["child_ids"][:1]
                contacts = odoo_call(models, uid, "res.partner", "read",
                    [target_ids], {"fields": ["name", "email"]})
                if contacts:
                    v["contact_name"] = contacts[0].get("name", "")
                    if not v.get("email"):
                        v["email"] = contacts[0].get("email", "")
        return {v["name"]: v for v in vendors}, None
    except Exception as e:
        return {}, str(e)

@st.cache_data(ttl=300, show_spinner="Loading payment terms...")
def fetch_payment_terms():
    """Returns ({display_name: id}, error). Cached 5 min — these change rarely."""
    try:
        uid, models = get_odoo_connection()
        terms = odoo_call(models, uid, "account.payment.term", "search_read",
            [[("active", "=", True)]],
            {"fields": ["id", "name"], "order": "sequence asc, id asc"})
        return {t["name"]: t["id"] for t in terms}, None
    except Exception as e:
        return {}, str(e)

@st.cache_data(ttl=300, show_spinner="Loading incoterms...")
def fetch_incoterms():
    """Returns ({code: id}, error). Cached 5 min."""
    try:
        uid, models = get_odoo_connection()
        terms = odoo_call(models, uid, "account.incoterms", "search_read",
            [[("active", "=", True)]],
            {"fields": ["id", "code", "name"], "order": "code asc"})
        # Display "DDP — DELIVERED DUTY PAID" but key by code so user picks easily
        return {f"{t['code']} — {t['name']}": t["id"] for t in terms}, None
    except Exception as e:
        return {}, str(e)

@st.cache_data(ttl=300, show_spinner="Loading users...")
def fetch_users():
    """Returns ({display_name: id}, error). Internal users only.
    Used for fuzzy-matching the free-text requisitioner field to a user_id."""
    try:
        uid, models = get_odoo_connection()
        users = odoo_call(models, uid, "res.users", "search_read",
            [[("share", "=", False), ("active", "=", True)]],
            {"fields": ["id", "name", "login"], "order": "name asc"})
        # Build a map keyed by both name AND login email for fuzzy matching
        user_map = {}
        for u in users:
            user_map[u["name"].lower()] = u["id"]
            user_map[u["login"].lower()] = u["id"]
            # Also map first name only for casual matches like "Stephan"
            first = u["name"].split()[0].lower() if u["name"] else ""
            if first and first not in user_map:
                user_map[first] = u["id"]
        return user_map, None
    except Exception as e:
        return {}, str(e)

@st.cache_data(ttl=120, show_spinner="Loading products from Odoo...")
def fetch_products(po_type_key):
    """Purchasable Odoo products for this PO type, plus the generic fallback.

    Returns (products, error). Each product is a dict with:
    id, name, default_code, uom_name, is_storable, generic, label.
    The generic product (SWR Glass Panel / SWR Aluminium Extrusion) is the
    fallback until the real products are in Odoo.
    """
    try:
        uid, models = get_odoo_connection()
        categ_word = "Glass" if po_type_key == "Glass" else "Aluminum Profiles"
        generic_name = rules.GENERIC_PRODUCT_NAMES[po_type_key]
        rows = odoo_call(models, uid, "product.product", "search_read",
            [["|", ("name", "ilike", generic_name),
              "&", ("purchase_ok", "=", True),
                   ("categ_id.complete_name", "ilike", categ_word)]],
            {"fields": ["id", "name", "default_code", "uom_id", "is_storable"],
             "order": "default_code asc, name asc", "limit": 500})
        products = []
        for r in rows:
            prod = {
                "id": r["id"],
                "name": r.get("name") or "",
                "default_code": r.get("default_code") or "",
                "uom_name": r["uom_id"][1] if r.get("uom_id") else "",
                "is_storable": bool(r.get("is_storable")),
            }
            prod["generic"] = rules.is_generic(prod, po_type_key)
            code = f"[{prod['default_code']}] " if prod["default_code"] else ""
            tail = "  (generic, not real stock)" if prod["generic"] else f"  ({prod['uom_name']})"
            prod["label"] = f"{code}{prod['name']}{tail}"
            products.append(prod)
        # Two products can have the same name and code: make each label unique.
        seen = {}
        for prod in products:
            n = seen.get(prod["label"], 0)
            seen[prod["label"]] = n + 1
            if n:
                prod["label"] = f"{prod['label']}  #{prod['id']}"
        return products, None
    except Exception as e:
        return [], str(e)

# ─────────────────────────────────────────────────────────────
# FILE PARSERS
# ─────────────────────────────────────────────────────────────
def parse_glass_file(file_bytes):
    """Parse SWR_Glass xlsx — returns list of glass line items."""
    try:
        # Glass file has header info in first 12 rows, data starts at row 12 (0-indexed)
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Glass", header=12)
        # If Tag column not found, try finding it dynamically
        if "Tag" not in df.columns:
            raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Glass", header=None)
            for i, row in raw.iterrows():
                if "Tag" in str(row.values):
                    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Glass", header=i)
                    break
        # Drop totals row
        df = df[df["Tag"].astype(str).str.strip() != "Totals"].copy()
        df = df.dropna(subset=["Tag"])
        # Glass type (GL-1, GL-2, ...) is in newer cutlist files only.
        # Older files have no such column, so glass_type stays "".
        gt_col = next((c for c in df.columns
                       if str(c).strip().lower() == "glass type"), None)
        lines = []
        for _, row in df.iterrows():
            try:
                glass_type = ""
                if gt_col is not None and pd.notna(row.get(gt_col)):
                    glass_type = str(row.get(gt_col)).strip()
                area_each  = float(row.get("Area Each (ft²)", 0) or 0)
                qty        = int(float(row.get("Qty", 0) or 0))
                area_total = float(row.get("Area Total (ft²)", 0) or 0)
                w_16 = str(row.get("Glass Width (1/16)", "")).strip()
                h_16 = str(row.get("Glass Height (1/16)", "")).strip()
                size_str = f'{w_16}" x {h_16}"' if w_16 and h_16 else ""
                if qty > 0:
                    lines.append({
                        "tag":        str(row.get("Tag", "")),
                        "glass_type": glass_type,
                        "size_str":   size_str,
                        "area_each":  round(area_each, 2),
                        "qty":        qty,
                        "area_total": round(area_total, 2),
                        "description": "",
                    })
            except Exception:
                continue
        return lines, None
    except Exception as e:
        return [], str(e)

def parse_aggcutonly_file(file_bytes):
    """Parse SWR_AggCutOnly xlsx — returns list of aluminium cut items."""
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="AggCutOnly")
        if "Finished Length in" not in df.columns:
            # Try to find header row
            for i, row in df.iterrows():
                if "Finished Length in" in str(row.values):
                    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="AggCutOnly",
                                       header=i)
                    break
        df = df.dropna(subset=["Finished Length in"])
        lines = []
        for _, row in df.iterrows():
            try:
                length   = float(row.get("Finished Length in", 0) or 0)
                total_qty = int(float(row.get("Total QTY", 0) or 0))
                part_num  = str(row.get("Part #", "")).strip()
                if length > 0 and total_qty > 0:
                    lines.append({
                        "length_in":   round(length, 3),
                        "length_ft":   round(length / 12, 3),
                        "part_number": part_num,
                        "qty":         total_qty,
                        "description": f"{part_num} — {length:.3f}\" extrusion",
                    })
            except Exception:
                continue
        return lines, None
    except Exception as e:
        return [], str(e)

def parse_optimizer_file(file_bytes):
    """Parse cutting_list optimizer Excel — returns line items grouped by (Part Number, Stock Length)."""
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=None)

        # Find the DETAILED CUT LIST header row
        detail_row = None
        for i, row in df.iterrows():
            if str(row.iloc[0]).strip() == "DETAILED CUT LIST":
                detail_row = i
                break
        if detail_row is None:
            return [], "Could not find 'DETAILED CUT LIST' section in optimizer file."

        # The column header row is right after
        col_row = detail_row + 1
        df_detail = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=col_row)

        # Keep only rows that have Stock Length (ft) — these are the stock piece rows
        df_detail = df_detail.dropna(subset=["Stock Length (ft)"])
        df_detail = df_detail[pd.to_numeric(df_detail["Stock Length (ft)"], errors="coerce").notna()]

        # Group by (Part Number, Stock Length (ft)) and count rows = number of stock pieces
        df_detail["Stock Length (ft)"] = pd.to_numeric(df_detail["Stock Length (ft)"], errors="coerce")
        df_detail["Part Number"] = df_detail["Part Number"].fillna("").astype(str).str.strip()

        grouped = (
            df_detail.groupby(["Part Number", "Stock Length (ft)"])
            .size()
            .reset_index(name="qty")
        )

        lines = []
        for _, row in grouped.iterrows():
            length_ft = float(row["Stock Length (ft)"])
            length_in = round(length_ft * 12, 3)
            lines.append({
                "profile":     str(row["Part Number"]),
                "length_in":   length_in,
                "length_ft":   round(length_ft, 3),
                "qty":         int(row["qty"]),
                # PO-specific fields — user fills in app
                "alloy":       "6063-T6",
                "finish":      "",
                "lead_time":   "3-4 weeks",
                "tooling":     "",
                "die_setup":   "",
                "unit_price":  0.0,
                "um":          "Pc",
            })
        return lines, None
    except Exception as e:
        return [], str(e)


# ─────────────────────────────────────────────────────────────
# ALUMINUM PO DOCX GENERATOR
# ─────────────────────────────────────────────────────────────
def _add_logo(cell, logo_path, width_in=None):
    """Place the logo in a table cell, centered horizontally and vertically."""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    p.add_run().add_picture(logo_path, width=Inches(width_in or LOGO_WIDTH_IN))


def generate_aluminum_po_docx(
    vendor_name, vendor_contact, vendor_email, vendor_address,
    ship_to_lines, job_number, job_location,
    po_date, po_number, requisitioner,
    shipped_via, fob_point, terms,
    al_lines,          # list of dicts with profile, length_in, qty, alloy, finish, lead_time, tooling, die_setup, unit_price, um
    packaging_note="Fully Corrugated Bundles, Paper Layer Separation",
    logo_path=None,
):
    doc = DocxDocument()
    for section in doc.sections:
        section.top_margin    = Cm(1.27)
        section.bottom_margin = Cm(1.27)
        section.left_margin   = Cm(1.27)
        section.right_margin  = Cm(1.27)

    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(9)

    subtotal   = sum(l["unit_price"] * l["qty"] for l in al_lines)
    grand_total = subtotal  # shipping/tax added below if non-zero

    # 3 tables, like the glass PO: header (9 cols) / PO metadata (6 cols) / line items (9 cols).
    # The metadata gets its own table so its 6 fields are not forced into the line-item columns.
    # All 3 tables sum to 16.8cm so their edges line up.
    # Columns: Item# | Profile# | Alloy | Finish | Length(in) | Lead Time | Pcs | Unit Price | Total
    col_widths = [Cm(1.2), Cm(1.8), Cm(1.8), Cm(2.8), Cm(1.6), Cm(2.2), Cm(1.2), Cm(2.0), Cm(2.2)]

    def _new_table(widths):
        t = doc.add_table(rows=0, cols=len(widths))
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.style = "Table Grid"
        t.autofit = False
        for i, w in enumerate(widths):
            t.columns[i].width = w
        return t

    table = _new_table(col_widths)

    # ── Logo + PURCHASE ORDER ──
    row = table.add_row()
    c = row.cells[0]; c.merge(row.cells[3])
    if logo_path and os.path.exists(logo_path):
        _add_logo(c, logo_path)
    c = row.cells[4]; c.merge(row.cells[8])
    c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run("PURCHASE ORDER")
    run.bold = True; run.font.size = Pt(16); run.font.small_caps = True

    # ── TO / SHIP TO / PO NUMBER ──
    row = table.add_row()
    c = row.cells[0]; c.merge(row.cells[2])
    p = c.paragraphs[0]; run = p.add_run("TO:"); run.bold = True; run.font.small_caps = True
    if vendor_contact: c.add_paragraph().add_run(vendor_contact)
    c.add_paragraph().add_run(vendor_name)
    if vendor_address:
        for ln in vendor_address.split("\n"):
            if ln.strip(): c.add_paragraph().add_run(ln)
    c.add_paragraph()
    p = c.add_paragraph(); p.add_run(f"JOB NO.: {job_number}").bold = True
    if job_location:
        c.add_paragraph().add_run(f"JOB LOCATION: {job_location}").bold = True

    c = row.cells[3]; c.merge(row.cells[5])
    p = c.paragraphs[0]; run = p.add_run("SHIP TO:"); run.bold = True; run.font.small_caps = True
    for ln in ship_to_lines: c.add_paragraph().add_run(ln)

    c = row.cells[6]; c.merge(row.cells[8])
    p = c.paragraphs[0]; run = p.add_run("P.O. Number:"); run.bold = True
    c.add_paragraph().add_run(po_number)
    c.add_paragraph().add_run(
        "The P.O. number must appear on all related correspondence, shipping papers, and invoices."
    ).font.size = Pt(7)

    doc.add_paragraph()  # spacer

    # ── PO metadata table (own widths; sums to 16.8cm) ──
    meta_widths = [Cm(2.4), Cm(3.0), Cm(3.0), Cm(2.2), Cm(2.6), Cm(3.6)]
    meta = _new_table(meta_widths)
    row = meta.add_row()
    for i, h in enumerate(["P.O. DATE", "P.O. NUMBER", "REQUISITIONER",
                           "SHIPPED VIA", "FREIGHT TERMS", "TERMS"]):
        row.cells[i].width = meta_widths[i]
        p = row.cells[i].paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h); run.bold = True; run.font.size = Pt(7); run.font.small_caps = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        _add_shading(row.cells[i], BRAND_TEAL)

    row = meta.add_row()
    for i, v in enumerate([po_date, po_number, requisitioner, shipped_via, fob_point, terms]):
        row.cells[i].width = meta_widths[i]
        p = row.cells[i].paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run(str(v)).font.size = Pt(9)

    doc.add_paragraph()  # spacer

    # ── Line items table (same 9 columns as the header table) ──
    table = _new_table(col_widths)

    # ── Line item headers ──
    headers = ["ITEM\n#", "PROFILE\nNUMBER", "ALLOY/\nTEMPER", "FINISH",
               "LENGTH\n(INCHES)", "APPROX.\nLEAD TIME", "PCS",
               "UNIT\nPRICE/PC", "TOTAL"]
    row = table.add_row()
    for i, h in enumerate(headers):
        row.cells[i].width = col_widths[i]
        p = row.cells[i].paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h); run.bold = True; run.font.size = Pt(7.5); run.font.small_caps = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        _add_shading(row.cells[i], BRAND_TEAL)

    # ── Line items ──
    for idx, line in enumerate(al_lines):
        row = table.add_row()
        line_total = line["unit_price"] * line["qty"]
        unit_price_str = f"${line['unit_price']:,.2f}" if line["unit_price"] else "-"
        tooling_str    = f"${line['tooling']:,.2f}" if line.get("tooling") and line["tooling"] else "-"
        vals = [
            str(idx + 1),
            line["profile"],
            line.get("alloy", "6063-T6"),
            line.get("finish", ""),
            str(line["length_in"]),
            line.get("lead_time", ""),
            str(line["qty"]),
            unit_price_str,
            f"${line_total:,.2f}" if line_total else "-",
        ]
        for i, v in enumerate(vals):
            row.cells[i].width = col_widths[i]
            p = row.cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if i == 3: p.alignment = WD_ALIGN_PARAGRAPH.LEFT   # Finish left-aligned
            if i == 8: p.alignment = WD_ALIGN_PARAGRAPH.RIGHT  # Total right-aligned
            p.add_run(v).font.size = Pt(9)

    # ── Empty buffer rows ──
    for _ in range(3):
        row = table.add_row()
        row.cells[0].merge(row.cells[8])

    # ── Packing + Subtotal ──
    row = table.add_row()
    c = row.cells[0]; c.merge(row.cells[5])
    run = c.paragraphs[0].add_run(f"Packing: {packaging_note}")
    run.underline = True; run.font.size = Pt(9)
    c = row.cells[6]; c.merge(row.cells[7])
    c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    c.paragraphs[0].add_run("SUBTOTAL").font.size = Pt(9)
    row.cells[8].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = row.cells[8].paragraphs[0].add_run(f"${subtotal:,.2f}"); run.bold = True; run.font.size = Pt(9)

    def _cost_row_al(label, amount_str):
        r = table.add_row()
        r.cells[0].merge(r.cells[5])
        # Instructions in left cell (first cost row only)
        c2 = r.cells[6]; c2.merge(r.cells[7])
        c2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        c2.paragraphs[0].add_run(label).font.size = Pt(9)
        r.cells[8].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r.cells[8].paragraphs[0].add_run(amount_str).font.size = Pt(9)

    _cost_row_al("SALES TAX", "")
    _cost_row_al("SHIPPING AND HANDLING", "")

    # ── Terms row ──
    row = table.add_row()
    c = row.cells[0]; c.merge(row.cells[5])
    for ln in [
        "1. Enter this order in accordance with the prices, terms, delivery method, and specifications listed in this purchase order.",
        "2. Please notify us immediately if you are unable to ship as specified.",
        "3. Send Invoices to accounts@inovues.com",
        "4. Send all correspondence to:",
        "   INOVUES, INC.",
        "   2700 Post Oak Blvd, 2100, Houston, TX 77056",
        "   (833) 466-8837 (INO-VUES)",
        "   info@inovues.com",
    ]:
        c.add_paragraph().add_run(ln).font.size = Pt(7)
    c2 = row.cells[6]; c2.merge(row.cells[7])
    c2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    c2.paragraphs[0].add_run("TOTAL").bold = True
    row.cells[8].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = row.cells[8].paragraphs[0].add_run(f"${subtotal:,.2f}"); run.bold = True; run.font.size = Pt(10)

    # ── Authorized by ──
    row = table.add_row()
    row.cells[0].merge(row.cells[4])
    c = row.cells[5]; c.merge(row.cells[7])
    c.paragraphs[0].add_run("Authorized by _____________________").font.size = Pt(9)
    row.cells[8].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    row.cells[8].paragraphs[0].add_run(po_date).font.size = Pt(9)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ─────────────────────────────────────────────────────────────
# PO DOCX GENERATOR (reused from SWR03111.py)
# ─────────────────────────────────────────────────────────────
def _add_shading(cell, color):
    tcPr = cell._element.get_or_add_tcPr()
    shd  = tcPr.makeelement(qn("w:shd"), {qn("w:fill"): color, qn("w:val"): "clear"})
    tcPr.append(shd)

def generate_po_docx(
    vendor_name, vendor_contact, vendor_email, vendor_address,
    ship_to_lines, job_number, job_location,
    po_date, po_number, requisitioner,
    lead_time, shipped_via, fob_point, terms,
    glass_lines,
    price_per_sqft, packaging_cost, shipping_cost, sales_tax, other_cost,
    packaging_note="Non-returnable boxed crate/rack",
    logo_path=None,
    unit_label="ft²",
    price_label="Price/ft²",
):
    """Glass PO generator — 3-table structure to prevent column-width fighting.

    Tables:
      1. Header table (4 cols)  — logo + PURCHASE ORDER title + TO/SHIP TO/BILL TO
      2. Metadata table (7 cols) — PO Date / Number / Requisitioner / Lead Time /
         Shipped Via / Freight Terms / Terms
      3. Line items table (8 cols) — Item # / Description / Unit Size / QTY /
         Area Each / Total ft² / Unit Price / Total, plus totals rows

    Line item math:
      Unit Price = price_per_sqft × area_each (per piece)
      Total per row = Unit Price × QTY (== price_per_sqft × area_total)
    """
    doc = DocxDocument()
    for section in doc.sections:
        section.top_margin    = Cm(1.27)
        section.bottom_margin = Cm(1.27)
        section.left_margin   = Cm(1.27)
        section.right_margin  = Cm(1.27)

    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(9)

    subtotal = sum(line["area_total"] * price_per_sqft for line in glass_lines)
    total    = subtotal + (sales_tax or 0) + (packaging_cost or 0) + (shipping_cost or 0) + (other_cost or 0)

    # Disable autofit globally for a moment so our cell widths stick
    def _disable_autofit(tbl):
        tbl.autofit = False
        tbl.allow_autofit = False

    # ─────────────────────────────────────────────────────────
    # TABLE 1 — Header (logo + title + TO/SHIP TO/BILL TO)
    # 4 columns of equal width (4.4cm each, sums to 17.6cm; all 3 tables must sum to 17.6cm)
    # ─────────────────────────────────────────────────────────
    t1 = doc.add_table(rows=0, cols=4)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    t1.style = "Table Grid"
    _disable_autofit(t1)
    t1_widths = [Cm(4.4), Cm(4.4), Cm(4.4), Cm(4.4)]
    for i, w in enumerate(t1_widths):
        t1.columns[i].width = w

    # Row 1: logo (cols 0-1) + PURCHASE ORDER title (cols 2-3)
    row = t1.add_row()
    for i, w in enumerate(t1_widths):
        row.cells[i].width = w
    c = row.cells[0]; c.merge(row.cells[1])
    if logo_path and os.path.exists(logo_path):
        _add_logo(c, logo_path)
    c = row.cells[2]; c.merge(row.cells[3])
    c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run("PURCHASE ORDER")
    run.bold = True; run.font.size = Pt(16); run.font.small_caps = True

    # Row 2: TO (cols 0-1) + SHIP TO (col 2) + BILL TO (col 3)
    row = t1.add_row()
    for i, w in enumerate(t1_widths):
        row.cells[i].width = w
    c = row.cells[0]; c.merge(row.cells[1])
    p = c.paragraphs[0]; run = p.add_run("TO:"); run.bold = True; run.font.small_caps = True
    c.add_paragraph().add_run(vendor_name)
    if vendor_contact: c.add_paragraph().add_run(f"Attn: {vendor_contact}")
    if vendor_email:   c.add_paragraph().add_run(vendor_email)
    if vendor_address:
        for ln in vendor_address.split("\n"):
            if ln.strip(): c.add_paragraph().add_run(ln)
    c.add_paragraph()
    p = c.add_paragraph(); run = p.add_run(f"JOB NO.: {job_number}"); run.bold = True
    if job_location:
        p = c.add_paragraph(); run = p.add_run(f"JOB LOCATION: {job_location}"); run.bold = True

    c = row.cells[2]
    p = c.paragraphs[0]; run = p.add_run("SHIP TO:"); run.bold = True; run.font.small_caps = True
    for ln in ship_to_lines: c.add_paragraph().add_run(ln)

    c = row.cells[3]
    p = c.paragraphs[0]; run = p.add_run("BILL TO:"); run.bold = True; run.font.small_caps = True
    for ln in ["INOVUES, INC.", "2700 Post Oak Blvd., 2100", "Houston, TX 77056",
               "accounts@inovues.com", "(833) 466-8837 (INO-VUES)"]:
        c.add_paragraph().add_run(ln)

    # Small spacer between tables
    doc.add_paragraph()

    # ─────────────────────────────────────────────────────────
    # TABLE 2 — PO Metadata (7 cols, evenly distributed)
    # Each column ~2.47cm; tweaked slightly by content length.
    # ─────────────────────────────────────────────────────────
    t2 = doc.add_table(rows=0, cols=7)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    t2.style = "Table Grid"
    _disable_autofit(t2)
    # PO DATE | PO NUMBER | REQUISITIONER | LEAD TIME | SHIPPED VIA | FREIGHT TERMS | TERMS
    t2_widths = [Cm(2.2), Cm(2.7), Cm(3.0), Cm(2.3), Cm(2.4), Cm(2.5), Cm(2.5)]  # sums to 17.6cm
    for i, w in enumerate(t2_widths):
        t2.columns[i].width = w

    meta_labels = ["P.O. DATE", "P.O. NUMBER", "REQUISITIONER", "LEAD TIME",
                   "SHIPPED VIA", "FREIGHT TERMS", "TERMS"]
    meta_values = [po_date, po_number, requisitioner, lead_time,
                   shipped_via, fob_point, terms]

    # Header row
    row = t2.add_row()
    for i, h in enumerate(meta_labels):
        row.cells[i].width = t2_widths[i]
        p = row.cells[i].paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h); run.bold = True; run.font.size = Pt(7); run.font.small_caps = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        _add_shading(row.cells[i], BRAND_TEAL)

    # Value row
    row = t2.add_row()
    for i, v in enumerate(meta_values):
        row.cells[i].width = t2_widths[i]
        p = row.cells[i].paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run(str(v)).font.size = Pt(9)

    # Spacer
    doc.add_paragraph()

    # ─────────────────────────────────────────────────────────
    # TABLE 3 — Line Items (8 cols) + totals/footer rows
    # Item # | Description | Unit Size | QTY | Area Each | Total ft² | Unit Price | Total
    # ─────────────────────────────────────────────────────────
    NCOLS = 8
    LAST = NCOLS - 1
    t3 = doc.add_table(rows=0, cols=NCOLS)
    t3.alignment = WD_TABLE_ALIGNMENT.CENTER
    t3.style = "Table Grid"
    _disable_autofit(t3)
    col_widths = [Cm(1.2), Cm(3.5), Cm(3.0), Cm(1.2), Cm(1.9), Cm(1.9), Cm(2.3), Cm(2.6)]  # sums to 17.6cm
    for i, w in enumerate(col_widths):
        t3.columns[i].width = w

    # Header row
    line_headers = [
        "ITEM\n#", "DESCRIPTION", "UNIT SIZE", "QTY",
        f"AREA EACH ({unit_label})", f"TOTAL ({unit_label})",
        "UNIT PRICE", "TOTAL",
    ]
    row = t3.add_row()
    for i, h in enumerate(line_headers):
        row.cells[i].width = col_widths[i]
        p = row.cells[i].paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h); run.bold = True; run.font.size = Pt(8); run.font.small_caps = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        _add_shading(row.cells[i], BRAND_TEAL)

    # Line items
    align_map = {
        0: WD_ALIGN_PARAGRAPH.CENTER,   # Item#
        3: WD_ALIGN_PARAGRAPH.CENTER,   # QTY
        4: WD_ALIGN_PARAGRAPH.RIGHT,    # Area Each
        5: WD_ALIGN_PARAGRAPH.RIGHT,    # Total ft²
        6: WD_ALIGN_PARAGRAPH.RIGHT,    # Unit Price
        7: WD_ALIGN_PARAGRAPH.RIGHT,    # Total
    }
    for idx, line in enumerate(glass_lines):
        row = t3.add_row()
        unit_price = line["area_each"] * price_per_sqft
        line_total = line["area_total"] * price_per_sqft
        vals = [
            str(idx + 1),
            line.get("description", ""),
            line.get("size_str", ""),
            str(line["qty"]),
            f"{line['area_each']:.3f}",
            f"{line['area_total']:.3f}",
            f"${unit_price:,.2f}",
            f"${line_total:,.2f}",
        ]
        for i, v in enumerate(vals):
            row.cells[i].width = col_widths[i]
            p = row.cells[i].paragraphs[0]
            if i in align_map:
                p.alignment = align_map[i]
            p.add_run(v).font.size = Pt(9)

    # Totals rows: cols 0-5 = label/notes area, col 6 = label, col 7 = amount
    def _totals_row(label, amount, label_bold=False, amount_bold=False, font_size=9):
        r = t3.add_row()
        for i, w in enumerate(col_widths):
            r.cells[i].width = w
        left = r.cells[0]; left.merge(r.cells[5])
        lab = r.cells[6]; lab.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run_l = lab.paragraphs[0].add_run(label)
        run_l.font.size = Pt(font_size); run_l.bold = label_bold
        amt = r.cells[LAST]; amt.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run_a = amt.paragraphs[0].add_run(f"${amount:,.2f}" if amount else "")
        run_a.font.size = Pt(font_size); run_a.bold = amount_bold
        return r, left

    # Subtotal row — also carries the packaging note on the left
    r, left = _totals_row("SUBTOTAL", subtotal, label_bold=False, amount_bold=True)
    run = left.paragraphs[0].add_run(f"Packaging: {packaging_note}")
    run.underline = True; run.font.size = Pt(9)

    _totals_row("SALES TAX", sales_tax)
    _totals_row("PACKAGING", packaging_cost)

    # Shipping row — also carries the order-instructions block on the left
    r, left = _totals_row("SHIPPING & HANDLING", shipping_cost)
    for ln in [
        "1. Enter this order in accordance with the prices, terms, delivery method, and specifications listed in this purchase order.",
        "2. Please notify us immediately if you are unable to ship as specified.",
        "3. Send all correspondence to:",
        "   INOVUES, INC.",
        "   2700 Post Oak Blvd, 2100, Houston, TX 77056",
        "   (833) 466-8837 (INO-VUES)",
        "   accounts@inovues.com",
    ]:
        left.add_paragraph().add_run(ln).font.size = Pt(7)

    _totals_row("OTHER", other_cost)
    _totals_row("TOTAL", total, label_bold=True, amount_bold=True, font_size=10)

    # Signature row
    row = t3.add_row()
    for i, w in enumerate(col_widths):
        row.cells[i].width = w
    row.cells[0].merge(row.cells[5])
    sig = row.cells[6]
    run = sig.paragraphs[0].add_run("Authorized by _____________________")
    run.font.size = Pt(9)
    row.cells[LAST].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    row.cells[LAST].paragraphs[0].add_run(po_date).font.size = Pt(9)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ─────────────────────────────────────────────────────────────
# ODOO PO CREATION
# ─────────────────────────────────────────────────────────────
def create_odoo_po(vendor, po_lines, price_per_unit, project_number, project_name,
                    po_number, requisitioner, subtotal, grand_total, po_buf,
                    po_type, po_date_str,
                    payment_term_id=None, incoterm_id=None,
                    date_planned=None, project_id=None, products=None):
    """Create a draft purchase.order in Odoo.

    Optional fields:
      payment_term_id  int, account.payment.term ID. From the Terms dropdown.
      incoterm_id      int, account.incoterms ID. From the Freight Terms dropdown.
      date_planned     datetime.date or datetime.datetime. From the date picker.
      project_id       int, project.project ID. Links the order to the project.
      products         list from fetch_products(). Each line in po_lines can
                       carry "product_id"; a line with none uses the generic
                       product.

    Odoo Workflow rule 3 (see odoo_rules.py): when the order has a project,
    Odoo puts the project tag on every line. This function removes the tag
    from lines for stock parts, because their cost must reach the project
    only through the manufacturing order.

    The 'requisitioner' free text is matched against active internal users;
    when a match is found, user_id is set on the PO.

    Returns (po_id, po_name, info). info has project_linked, tags_cleared,
    generic_lines.
    """
    uid, models = get_odoo_connection()

    def oc(model, method, args, kwargs={}):
        return odoo_call(models, uid, model, method, args, kwargs)

    products = products or []
    products_by_id = {p["id"]: p for p in products}
    generic = next((p for p in products if p.get("generic")), None)
    fallback_id = generic["id"] if generic else None

    # A line with no product needs the generic product. Fail loudly when it
    # is not in Odoo, rather than put a random product on the line.
    if any(not l.get("product_id") for l in po_lines) and not fallback_id:
        product_name = rules.GENERIC_PRODUCT_NAMES[po_type]
        product_ids = oc("product.product", "search", [[("name", "ilike", product_name)]])
        if not product_ids:
            raise Exception(
                f"One or more lines have no Odoo product, and the generic "
                f"product '{product_name}' is not in Odoo. "
                f"Select a real product for each line in Step 2."
            )
        fallback_id = product_ids[0]

    order_lines = rules.build_order_lines(
        po_type, po_lines, price_per_unit, products_by_id, fallback_id)
    if not order_lines:
        raise Exception("No line has a quantity above 0. Nothing to order.")
    generic_ids = {fallback_id} | {p["id"] for p in products if p.get("generic")}
    generic_lines = sum(1 for _, _, v in order_lines if v["product_id"] in generic_ids)

    # Build the PO payload, only including optional fields when provided
    po_vals = {
        "partner_id":  vendor["id"],
        "origin":      f"{project_number} / {po_type} PO",
        "order_line":  order_lines,
    }
    if project_id:
        po_vals["project_id"] = int(project_id)
    if payment_term_id:
        po_vals["payment_term_id"] = payment_term_id
    if incoterm_id:
        po_vals["incoterm_id"] = incoterm_id
    if date_planned:
        # Odoo expects "YYYY-MM-DD HH:MM:SS"
        if hasattr(date_planned, "strftime"):
            po_vals["date_planned"] = date_planned.strftime("%Y-%m-%d 12:00:00") \
                if not hasattr(date_planned, "hour") else date_planned.strftime("%Y-%m-%d %H:%M:%S")
        else:
            po_vals["date_planned"] = str(date_planned)

    # Fuzzy-match requisitioner free-text to a real Odoo user
    user_map, _ = fetch_users()
    if requisitioner and user_map:
        matched_uid = user_map.get(requisitioner.strip().lower())
        if matched_uid:
            po_vals["user_id"] = matched_uid

    po_id = oc("purchase.order", "create", [po_vals])
    if isinstance(po_id, (list, tuple)):
        po_id = po_id[0]
    po_data = oc("purchase.order", "read", [[po_id]], {"fields": ["name"]})
    po_name = po_data[0]["name"] if po_data else f"ID {po_id}"

    # Rule 3: remove the project tag from stock lines.
    tags_cleared = 0
    if project_id:
        odoo_lines = oc("purchase.order.line", "search_read",
                        [[("order_id", "=", po_id)]],
                        {"fields": ["id", "product_id", "analytic_distribution"]})
        line_product_ids = sorted({l["product_id"][0] for l in odoo_lines if l.get("product_id")})
        storable = {}
        if line_product_ids:
            for p in oc("product.product", "read", [line_product_ids], {"fields": ["is_storable"]}):
                storable[p["id"]] = bool(p.get("is_storable"))
        to_clear = rules.line_ids_to_clear(odoo_lines, storable)
        if to_clear:
            oc("purchase.order.line", "write", [to_clear, {"analytic_distribution": False}])
            tags_cleared = len(to_clear)

    # Attach docx
    oc("ir.attachment", "create", [{
        "name":      f"INOVUES_PO_{po_number}_{po_type}_{datetime.now().strftime('%Y%m%d')}.docx",
        "type":      "binary",
        "datas":     base64.b64encode(po_buf.getvalue()).decode("utf-8"),
        "res_model": "purchase.order",
        "res_id":    po_id,
        "mimetype":  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    }])

    # Chatter message
    project_note = (f"Odoo project: linked. Project tag removed from {tags_cleared} stock line(s) (rule 3)."
                    if project_id else "Odoo project: NOT linked (no project selected).")
    generic_note = (f"<br/>Generic product on {generic_lines} line(s): the receipt will not add to real stock."
                    if generic_lines else "")
    oc("purchase.order", "message_post", [[po_id]], {
        "body": (
            f"<b>📝 PO from INOVUES PO Generator</b><br/>"
            f"Project: {project_name}<br/>"
            f"{project_note}{generic_note}<br/>"
            f"PO Number: {po_number}<br/>"
            f"Requisitioner: {requisitioner}<br/>"
            f"Type: {po_type}<br/>"
            f"Vendor: {vendor['name']}<br/>"
            f"Subtotal: ${subtotal:,.2f} | Total: ${grand_total:,.2f}<br/>"
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        ),
        "message_type": "comment",
        "subtype_xmlid": "mail.mt_comment",
    })
    info = {"project_linked": bool(project_id), "tags_cleared": tags_cleared,
            "generic_lines": generic_lines}
    return po_id, po_name, info


def show_odoo_result(po_id, po_name, info):
    """Tell the person what Odoo now holds for the new draft order."""
    st.success(f"✅ Draft PO **{po_name}** created in Odoo!")
    st.info(f"🔗 {ODOO_URL}/web#id={po_id}&model=purchase.order&view_type=form")
    if info["project_linked"]:
        st.caption(f"Linked to the Odoo project. Project tag removed from "
                   f"{info['tags_cleared']} stock line(s), so the material cost "
                   f"reaches the project one time, at the manufacturing order.")
    else:
        st.warning("This order is NOT linked to an Odoo project. "
                   "Select the project in Step 1 for a job order.")
    if info["generic_lines"]:
        st.warning(f"{info['generic_lines']} line(s) use the generic product. "
                   f"The receipt will not add to real stock.")


def pick_products_ui(po_type_key, groups, candidates_for, key_prefix):
    """One Odoo product dropdown per group. Returns ({group: product or None}, products).

    groups          list of (group key, title for the dropdown)
    candidates_for  function(group key, real products) -> list of products that
                    fit this group. Exactly 1 = selected for the person.
                    2 or more = shown first, and the person must pick.

    The app never puts a guessed product on a line. With no clear fit it
    selects the generic product when Odoo has one, else it selects nothing
    and the order cannot be created until a person picks.
    """
    products, prod_err = fetch_products(po_type_key)
    if prod_err:
        st.warning(f"Could not load Odoo products: {prod_err}. "
                   f"The generic product will be used.")
    if not products:
        return {}, products
    by_label = {p["label"]: p for p in products}
    generic = next((p for p in products if p.get("generic")), None)
    real = [p for p in products if not p.get("generic")]
    chosen = {}
    for gkey, title in groups:
        fits = candidates_for(gkey, real) or []
        fit_ids = {p["id"] for p in fits}
        ordered = fits + [p for p in products if p["id"] not in fit_ids]
        labels = [p["label"] for p in ordered]
        default = fits[0] if len(fits) == 1 else generic
        if len(fits) > 1:
            default = None          # several finishes fit: a person must pick
        pick = st.selectbox(
            title, labels,
            index=labels.index(default["label"]) if default else None,
            placeholder=(f"Choose 1 of {len(fits)} that fit..." if len(fits) > 1
                         else "Choose a product..."),
            key=f"{key_prefix}_{gkey or 'none'}")
        chosen[gkey] = by_label[pick] if pick else None
    missing = [title for gkey, title in groups if chosen.get(gkey) is None]
    if missing:
        st.warning("No Odoo product selected for: " + ", ".join(missing)
                   + (". The generic product will be used." if generic else
                      ". Select one, or the order cannot be created in Odoo."))
    if any(p and p.get("generic") for p in chosen.values()):
        st.warning("The generic product is selected for one or more lines. "
                   "Odoo cannot add these lines to real stock. "
                   "Select the real product when it is in Odoo.")
    return chosen, products


# ─────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────
if os.path.exists(LOGO_PATH):
    st.image(LOGO_PATH, width=180)
st.title("📝 INOVUES PO Generator")
st.caption("Generate Glass and Aluminium Purchase Orders from Odoo project files")

# ── PO Type selector ──
po_type = st.radio("Select PO Type", ["Glass PO", "Aluminium PO"],
                   horizontal=True, label_visibility="collapsed")
po_type_key = "Glass" if po_type == "Glass PO" else "Aluminium"
file_keyword = "Glass" if po_type_key == "Glass" else "cutting_list"

st.divider()

# ── Step 1: Project + File selection ──
st.subheader("Step 1 — Select Project & File")

col_proj, col_file = st.columns([1, 2])

with col_proj:
    project_map, proj_err = fetch_projects()
    if proj_err:
        st.error(f"Could not load projects: {proj_err}")
        st.stop()

    selected_project = st.selectbox("Odoo Project", options=list(project_map.keys()),
                                    index=None, placeholder="Choose a project...")

attachments = []
selected_attachment = None
line_items = []

if selected_project:
    project_id = project_map[selected_project]
    attachments, att_err = fetch_project_attachments(project_id, po_type_key)

    if att_err:
        st.error(f"Could not fetch attachments: {att_err}")
    elif not attachments:
        st.warning("No Excel attachments found in this project's tasks.")
    else:
        # Only show files matching the file_keyword for this PO type
        relevant = [a for a in attachments if file_keyword.lower() in a["name"].lower()]

        if not relevant:
            with col_file:
                if po_type_key == "Glass":
                    st.warning(
                        "⚠️ **No Glass files found in this project.**\n\n"
                        "**Workflow to create one:**\n"
                        "1. Open the **SWR Cutlist** app\n"
                        "2. Run it on this project's window dimensions\n"
                        "3. Save to Odoo — it will attach `..._SWR_Glass_...xlsx` to an Engineering task\n"
                        "4. Come back here and refresh"
                    )
                else:
                    st.warning(
                        "⚠️ **No Cutting Optimizer files found in this project.**\n\n"
                        "**Workflow to create one:**\n"
                        "1. Open the **SWR Cutlist** app → save to Odoo (produces `..._SWR_AggCutOnly_...xlsx`)\n"
                        "2. Open the **Cutting Optimizer** app → load that AggCutOnly file → save to Odoo (produces `cutting_list_...xlsx`)\n"
                        "3. Come back here and refresh"
                    )
        else:
            with col_file:
                file_options = {f"{a['name']}  (task: {a['task_name']})": a for a in relevant}
                selected_file_key = st.selectbox(
                    f"Select {file_keyword} file",
                    options=list(file_options.keys()),
                    index=None,
                    placeholder="Choose a file..."
                )

            if selected_file_key:
                selected_attachment = file_options[selected_file_key]
                file_bytes = base64.b64decode(selected_attachment["datas"])

                if po_type_key == "Glass":
                    line_items, parse_err = parse_glass_file(file_bytes)
                else:
                    line_items, parse_err = parse_optimizer_file(file_bytes)

                if parse_err:
                    st.error(f"Could not parse file: {parse_err}")
                    line_items = []

# ── Manual upload fallback ──
upload_label = "cutting_list optimizer .xlsx file" if po_type_key == "Aluminium" else "Glass .xlsx file"
with st.expander("Or upload a file manually (fallback)", expanded=not selected_project):
    uploaded = st.file_uploader(f"Upload {upload_label}", type=["xlsx"])
    if uploaded:
        file_bytes = uploaded.read()
        if po_type_key == "Glass":
            line_items, parse_err = parse_glass_file(file_bytes)
        else:
            line_items, parse_err = parse_optimizer_file(file_bytes)
        if parse_err:
            st.error(f"Could not parse file: {parse_err}")
            line_items = []

# ══════════════════════════════════════════════════════════════
# GLASS PO FLOW
# ══════════════════════════════════════════════════════════════
if line_items and po_type_key == "Glass":
    st.divider()
    st.subheader(f"Step 2 — Review & Edit Line Items ({len(line_items)} lines)")

    total_qty  = sum(l["qty"] for l in line_items)
    total_area = sum(l["area_total"] for l in line_items)
    st.write(f"**{total_qty} pieces | {total_area:.2f} ft² total**")

    # ── Descriptions ──
    # Newer cutlist files carry a "Glass Type" per line (GL-1, GL-2, ...).
    # Then we ask for ONE description per glass type, and every line of
    # that type gets it (e.g. GL-2 and GL-3 with different low-e surfaces).
    # Older files have no type: fall back to one default description
    # plus optional per-line overrides.
    glass_types = []
    for l in line_items:
        gt = l.get("glass_type", "")
        if gt and gt not in glass_types:
            glass_types.append(gt)

    if glass_types:
        st.write(f"Enter one description per glass type "
                 f"({len(glass_types)} type{'s' if len(glass_types) != 1 else ''} in this file):")
        type_desc = {}
        for gt in glass_types:
            gt_lines = [l for l in line_items if l.get("glass_type") == gt]
            tags = [l["tag"] for l in gt_lines]
            tag_txt = ", ".join(tags[:8]) + (f" +{len(tags) - 8} more" if len(tags) > 8 else "")
            pcs = sum(l["qty"] for l in gt_lines)
            type_desc[gt] = st.text_input(
                f"{gt}: {len(gt_lines)} line{'s' if len(gt_lines) != 1 else ''}, {pcs} pcs ({tag_txt})",
                placeholder=f"e.g. {gt}: 1\" IGU, 6mm low-E #2 + 12mm AS + 6mm clear",
                key=f"glass_desc_type_{gt}")
        untyped = [l for l in line_items if not l.get("glass_type")]
        if untyped:
            st.warning(f"{len(untyped)} line(s) have no Glass Type in the cutlist file: "
                       + ", ".join(l["tag"] for l in untyped))
            type_desc[""] = st.text_input("Description for lines with no glass type",
                                          key="glass_desc_type__none")
        for line in line_items:
            line["description"] = type_desc.get(line.get("glass_type", ""), "").strip()
        missing = [gt or "no type" for gt, d in type_desc.items() if not d.strip()]
        if missing:
            st.warning("No description yet for: " + ", ".join(missing))
    else:
        st.info("This file has no Glass Type column (made with an older SWR Cutlist). "
                "Enter one default description, and change single lines below if needed. "
                "Re-run the SWR Cutlist to get one box per glass type.")
        first_desc = st.text_input("Description for all lines (default)",
                                   placeholder="e.g. GT1 – 10mm Leadus VIG: 5Tlow-E+V+5T",
                                   key="glass_desc_default")
        with st.expander("Change the description for single lines", expanded=False):
            for i, line in enumerate(line_items):
                override = st.text_input(f"{line['tag']}  {line['size_str']}",
                                         placeholder=first_desc or "Same as default",
                                         key=f"glass_desc_line_{i}")
                line["description"] = override.strip() or first_desc

    # ── Odoo product per glass type (Odoo Workflow, Nov 2026) ──
    st.write("**Odoo product.** A real product lets the receipt add to the correct stock.")
    if glass_types:
        glass_groups = [(gt, f"Odoo product for {gt}") for gt in glass_types]
        if any(not l.get("glass_type") for l in line_items):
            glass_groups.append(("", "Odoo product for lines with no glass type"))
    else:
        glass_groups = [("", "Odoo product for all lines")]
    glass_chosen, glass_products = pick_products_ui(
        "Glass", glass_groups, lambda gkey, products: [], "glass_prod")
    for line in line_items:
        gkey = line.get("glass_type", "") if glass_types else ""
        line["product_id"] = (glass_chosen.get(gkey) or {}).get("id")

    preview_rows = []
    for l in line_items:
        row = {"Tag": l["tag"]}
        if glass_types:
            row["Glass Type"] = l.get("glass_type", "")
        row.update({
            "Description": l["description"], "Size": l["size_str"],
            "Area Each (ft²)": l["area_each"], "Qty": l["qty"],
            "Area Total (ft²)": l["area_total"],
        })
        preview_rows.append(row)
    st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Step 3 — Vendor & PO Details")

    vcol1, vcol2 = st.columns(2)
    with vcol1:
        vendor_map, vendor_err = fetch_vendors()
        if vendor_err:
            st.error(f"Could not load vendors: {vendor_err}")
        selected_vendor = st.selectbox("Select Vendor", options=list(vendor_map.keys()),
                                       index=None, placeholder="Choose a vendor...")
        if selected_vendor:
            v = vendor_map[selected_vendor]
            with st.expander("Vendor details", expanded=False):
                st.text(f"Email: {v.get('email','')}")
                st.text(f"Address: {v.get('full_address','')}")
    with vcol2:
        project_number = st.text_input("Project / PO Number", value="INO-")
        requisitioner  = st.text_input("Requisitioner", value="Stephan Ketterer")
        job_location   = st.text_input("Job Location", value="")

    dcol1, dcol2, dcol3 = st.columns(3)
    with dcol1:
        # Lead Time → date picker (writes to date_planned in Odoo)
        from datetime import timedelta
        lead_date = st.date_input("Expected Delivery Date",
                                  value=datetime.now().date() + timedelta(days=14),
                                  help="Becomes 'date_planned' on the Odoo PO.")
        lead_time = lead_date.strftime("%m/%d/%Y")  # human-readable for the .docx
        shipped_via = st.selectbox("Shipped Via", ["Air", "Ground"], index=1)
    with dcol2:
        # Freight Terms → dropdown of all 11 Odoo incoterms
        incoterm_map, incoterm_err = fetch_incoterms()
        if incoterm_err:
            st.warning(f"Could not load incoterms: {incoterm_err}")
        # Default to DAP (per Anas, Nov 2026; was DDP)
        incoterm_keys = list(incoterm_map.keys()) if incoterm_map else []
        dap_idx = next((i for i, k in enumerate(incoterm_keys) if k.startswith("DAP")), 0)
        fob_label = st.selectbox("Freight Terms", incoterm_keys,
                                  index=dap_idx if incoterm_keys else None,
                                  help="Becomes 'incoterm_id' on the Odoo PO.")
        fob_point = fob_label.split(" — ")[0] if fob_label else ""  # short code for .docx

        # Terms → dropdown of payment terms from Odoo
        payterm_map, payterm_err = fetch_payment_terms()
        if payterm_err:
            st.warning(f"Could not load payment terms: {payterm_err}")
        payterm_keys = list(payterm_map.keys()) if payterm_map else []
        # Default to "30 Days" (closest match to old "Net 30" default)
        net30_idx = next((i for i, k in enumerate(payterm_keys) if "30" in k.lower() and "balance" not in k.lower() and "10th" not in k.lower()), 0)
        terms = st.selectbox("Payment Terms", payterm_keys,
                              index=net30_idx if payterm_keys else None,
                              help="Becomes 'payment_term_id' on the Odoo PO.")
    with dcol3:
        price_per_unit = st.number_input("Price per ft² ($)", value=0.0, min_value=0.0,
                                          step=0.01, format="%.2f")

    ship_to_text = st.text_area("Ship To", value=SHIP_TO_DEFAULT, height=100)

    pcol1, pcol2 = st.columns(2)
    with pcol1:
        packaging_cost = st.number_input("Packaging ($)", value=0.0, min_value=0.0, format="%.2f")
        sales_tax      = st.number_input("Sales Tax ($)", value=0.0, min_value=0.0, format="%.2f")
        packaging_note = st.text_input("Packaging Note", value="Non-returnable boxed crate/rack")
    with pcol2:
        shipping_cost = st.number_input("Shipping ($)", value=0.0, min_value=0.0, format="%.2f")
        other_cost    = st.number_input("Other ($)", value=0.0, min_value=0.0, format="%.2f")

    subtotal    = sum(l["area_total"] for l in line_items) * price_per_unit
    grand_total = subtotal + sales_tax + packaging_cost + shipping_cost + other_cost
    st.write(f"**Subtotal: ${subtotal:,.2f}  |  Grand Total: ${grand_total:,.2f}**")

    st.divider()
    st.subheader("Step 4 — Generate PO")
    vendor_ready = selected_vendor is not None and vendor_map
    gcol1, gcol2 = st.columns(2)

    with gcol1:
        if st.button("📄 Generate PO (.docx)", disabled=not vendor_ready, key="glass_gen"):
            v = vendor_map[selected_vendor]
            po_buf = generate_po_docx(
                vendor_name=v["name"], vendor_contact=v.get("contact_name", ""),
                vendor_email=v.get("email", ""), vendor_address=v.get("full_address", ""),
                ship_to_lines=[ln for ln in ship_to_text.split("\n") if ln.strip()],
                job_number=project_number, job_location=job_location,
                po_date=datetime.now().strftime("%m/%d/%Y"), po_number=project_number,
                requisitioner=requisitioner, lead_time=lead_time, shipped_via=shipped_via,
                fob_point=fob_point, terms=terms, glass_lines=line_items,
                price_per_sqft=price_per_unit, packaging_cost=packaging_cost,
                shipping_cost=shipping_cost, sales_tax=sales_tax, other_cost=other_cost,
                packaging_note=packaging_note, unit_label="ft²",
                logo_path=LOGO_PATH,
            )
            st.session_state["po_buf"] = po_buf.getvalue()
            st.session_state["po_buf_name"] = f"INOVUES_PO_{project_number}_Glass_{datetime.now().strftime('%Y%m%d')}.docx"
            st.success("✅ PO document ready!")
        if "po_buf" in st.session_state:
            st.download_button("💾 Download PO .docx", data=st.session_state["po_buf"],
                               file_name=st.session_state.get("po_buf_name", "INOVUES_PO.docx"),
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    with gcol2:
        if st.button("📝 Create Draft PO in Odoo", type="primary", disabled=not vendor_ready, key="glass_odoo"):
            if "po_buf" not in st.session_state:
                st.warning("Generate the .docx first.")
            else:
                with st.spinner("Creating PO in Odoo..."):
                    try:
                        v = vendor_map[selected_vendor]
                        po_id, po_name, po_info = create_odoo_po(
                            vendor=v, po_lines=line_items, price_per_unit=price_per_unit,
                            project_number=project_number, project_name=selected_project or "",
                            po_number=project_number, requisitioner=requisitioner,
                            subtotal=subtotal, grand_total=grand_total,
                            po_buf=io.BytesIO(st.session_state["po_buf"]),
                            po_type="Glass", po_date_str=datetime.now().strftime("%m/%d/%Y"),
                            payment_term_id=payterm_map.get(terms) if terms else None,
                            incoterm_id=incoterm_map.get(fob_label) if fob_label else None,
                            date_planned=lead_date,
                            project_id=project_map.get(selected_project) if selected_project else None,
                            products=glass_products,
                        )
                        show_odoo_result(po_id, po_name, po_info)
                    except Exception as e:
                        st.error(f"❌ Error: {str(e)}")

# ══════════════════════════════════════════════════════════════
# ALUMINIUM PO FLOW
# ══════════════════════════════════════════════════════════════
elif line_items and po_type_key == "Aluminium":
    st.divider()
    st.subheader(f"Step 2 — Review & Edit Line Items ({len(line_items)} lines)")
    st.caption("Pre-filled from optimizer output. Adjust quantities and fill in pricing details.")

    # ── Per-line editable fields ──
    st.write("**Edit each line item:**")

    # Shared fields header
    sh1, sh2, sh3, sh4, sh5, sh6, sh7 = st.columns([1.5, 1.5, 2, 2, 1, 1.5, 1.5])
    sh1.markdown("**Profile #**"); sh2.markdown("**Length (in)**")
    sh3.markdown("**Alloy/Temper**"); sh4.markdown("**Finish**")
    sh5.markdown("**Qty**"); sh6.markdown("**Lead Time**"); sh7.markdown("**Unit Price ($)**")

    for i, line in enumerate(line_items):
        c1, c2, c3, c4, c5, c6, c7 = st.columns([1.5, 1.5, 2, 2, 1, 1.5, 1.5])
        line["profile"]    = c1.text_input("", value=line["profile"],    key=f"al_prof_{i}", label_visibility="collapsed")
        line["length_in"]  = c2.number_input("", value=float(line["length_in"]), min_value=0.0, step=1.0, key=f"al_len_{i}", label_visibility="collapsed")
        line["alloy"]      = c3.text_input("", value=line.get("alloy", "6063-T6"), key=f"al_alloy_{i}", label_visibility="collapsed")
        line["finish"]     = c4.text_input("", value=line.get("finish", ""),       key=f"al_finish_{i}", label_visibility="collapsed")
        line["qty"]        = c5.number_input("", value=int(line["qty"]), min_value=0, step=1, key=f"al_qty_{i}", label_visibility="collapsed")
        line["lead_time"]  = c6.text_input("", value=line.get("lead_time", "3-4 weeks"), key=f"al_lead_{i}", label_visibility="collapsed")
        line["unit_price"] = c7.number_input("", value=float(line.get("unit_price", 0.0)), min_value=0.0, step=0.01, format="%.2f", key=f"al_price_{i}", label_visibility="collapsed")

    # ── Odoo product per line (Odoo Workflow, Nov 2026) ──
    st.write("**Odoo product for each line.** The app proposes the product whose code fits the profile number.")
    al_groups = [(str(i), f"{l['profile']}  {l['length_in']}\"") for i, l in enumerate(line_items)]
    al_chosen, al_products = pick_products_ui(
        "Aluminium", al_groups,
        lambda gkey, products: rules.match_candidates(
            products, line_items[int(gkey)]["profile"]),
        "al_prod")
    for i, line in enumerate(line_items):
        line["product_id"] = (al_chosen.get(str(i)) or {}).get("id")

    total_pcs = sum(l["qty"] for l in line_items)
    subtotal_al = sum(l["unit_price"] * l["qty"] for l in line_items)
    st.write(f"**{total_pcs} total pcs | Subtotal: ${subtotal_al:,.2f}**")

    st.divider()
    st.subheader("Step 3 — Vendor & PO Details")

    vcol1, vcol2 = st.columns(2)
    with vcol1:
        vendor_map, vendor_err = fetch_vendors()
        if vendor_err:
            st.error(f"Could not load vendors: {vendor_err}")
        selected_vendor = st.selectbox("Select Vendor", options=list(vendor_map.keys()),
                                       index=None, placeholder="Choose a vendor...")
        if selected_vendor:
            v = vendor_map[selected_vendor]
            with st.expander("Vendor details", expanded=False):
                st.text(f"Email: {v.get('email','')}")
                st.text(f"Address: {v.get('full_address','')}")
    with vcol2:
        project_number = st.text_input("Project / PO Number", value="INO-", key="al_proj_num")
        requisitioner  = st.text_input("Requisitioner", value="Stephan Ketterer", key="al_req")
        job_location   = st.text_input("Job Location", value="", key="al_job_loc")

    dcol1, dcol2 = st.columns(2)
    with dcol1:
        from datetime import timedelta
        lead_date_al = st.date_input("Expected Delivery Date",
                                     value=datetime.now().date() + timedelta(days=28),
                                     key="al_lead_date",
                                     help="Becomes 'date_planned' on the Odoo PO.")
        lead_time_al = lead_date_al.strftime("%m/%d/%Y")
        shipped_via = st.selectbox("Shipped Via", ["Air", "Ground"], index=1, key="al_ship")
        # Freight Terms dropdown
        incoterm_map, incoterm_err = fetch_incoterms()
        if incoterm_err:
            st.warning(f"Could not load incoterms: {incoterm_err}")
        incoterm_keys = list(incoterm_map.keys()) if incoterm_map else []
        # Default to DAP, same as glass (per Anas, Nov 2026; was FOB)
        fob_idx = next((i for i, k in enumerate(incoterm_keys) if k.startswith("DAP")), 0)
        fob_label = st.selectbox("Freight Terms", incoterm_keys,
                                  index=fob_idx if incoterm_keys else None,
                                  key="al_fob",
                                  help="Becomes 'incoterm_id' on the Odoo PO.")
        fob_point = fob_label.split(" — ")[0] if fob_label else ""
    with dcol2:
        # Payment terms dropdown
        payterm_map, payterm_err = fetch_payment_terms()
        if payterm_err:
            st.warning(f"Could not load payment terms: {payterm_err}")
        payterm_keys = list(payterm_map.keys()) if payterm_map else []
        # Aluminium default was "Per Quote" — closest match in Odoo is "30 Days"
        net30_idx = next((i for i, k in enumerate(payterm_keys) if "30" in k.lower() and "balance" not in k.lower() and "10th" not in k.lower()), 0)
        terms = st.selectbox("Payment Terms", payterm_keys,
                              index=net30_idx if payterm_keys else None,
                              key="al_terms",
                              help="Becomes 'payment_term_id' on the Odoo PO.")
        packaging_note  = st.text_input("Packing Note", value="Fully Corrugated Bundles, Paper Layer Separation", key="al_pack_note")

    ship_to_text = st.text_area("Ship To", value=SHIP_TO_DEFAULT, height=100, key="al_ship_to")

    grand_total_al = subtotal_al
    st.write(f"**Grand Total: ${grand_total_al:,.2f}**")

    st.divider()
    st.subheader("Step 4 — Generate PO")
    vendor_ready = selected_vendor is not None and vendor_map
    gcol1, gcol2 = st.columns(2)

    with gcol1:
        if st.button("📄 Generate Aluminium PO (.docx)", disabled=not vendor_ready, key="al_gen"):
            v = vendor_map[selected_vendor]
            po_buf = generate_aluminum_po_docx(
                vendor_name=v["name"], vendor_contact=v.get("contact_name", ""),
                vendor_email=v.get("email", ""), vendor_address=v.get("full_address", ""),
                ship_to_lines=[ln for ln in ship_to_text.split("\n") if ln.strip()],
                job_number=project_number, job_location=job_location,
                po_date=datetime.now().strftime("%m/%d/%Y"), po_number=project_number,
                requisitioner=requisitioner, shipped_via=shipped_via,
                fob_point=fob_point, terms=terms,
                al_lines=line_items,
                packaging_note=packaging_note,
                logo_path=LOGO_PATH,
            )
            st.session_state["al_po_buf"] = po_buf.getvalue()
            st.session_state["al_po_buf_name"] = f"INOVUES_PO_{project_number}_Aluminium_{datetime.now().strftime('%Y%m%d')}.docx"
            st.success("✅ Aluminium PO document ready!")
        if "al_po_buf" in st.session_state:
            st.download_button("💾 Download Aluminium PO .docx",
                               data=st.session_state["al_po_buf"],
                               file_name=st.session_state.get("al_po_buf_name", "INOVUES_PO_Aluminium.docx"),
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               key="al_dl")

    with gcol2:
        if st.button("📝 Create Draft PO in Odoo", type="primary", disabled=not vendor_ready, key="al_odoo"):
            if "al_po_buf" not in st.session_state:
                st.warning("Generate the .docx first.")
            else:
                with st.spinner("Creating PO in Odoo..."):
                    try:
                        v = vendor_map[selected_vendor]
                        # Each line goes to Odoo with its own product, quantity and
                        # unit price (before Nov 2026 every line was sent at $1.00).
                        po_id, po_name, po_info = create_odoo_po(
                            vendor=v, po_lines=line_items, price_per_unit=None,
                            project_number=project_number, project_name=selected_project or "",
                            po_number=project_number, requisitioner=requisitioner,
                            subtotal=subtotal_al, grand_total=grand_total_al,
                            po_buf=io.BytesIO(st.session_state["al_po_buf"]),
                            po_type="Aluminium", po_date_str=datetime.now().strftime("%m/%d/%Y"),
                            payment_term_id=payterm_map.get(terms) if terms else None,
                            incoterm_id=incoterm_map.get(fob_label) if fob_label else None,
                            date_planned=lead_date_al,
                            project_id=project_map.get(selected_project) if selected_project else None,
                            products=al_products,
                        )
                        show_odoo_result(po_id, po_name, po_info)
                    except Exception as e:
                        st.error(f"❌ Error: {str(e)}")

else:
    if selected_project:
        st.info("Select a file above to continue.")
    else:
        st.info("Select an Odoo project to get started.")
