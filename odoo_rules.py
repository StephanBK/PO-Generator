"""Rules that make a PO Generator order fit the INOVUES Odoo workflow.

Pure functions only: no Streamlit and no Odoo connection in this file,
so that the rules can be tested without a server.

The rules come from "Odoo Workflow Blueprint v1" (Nov 2026):

  Rule 2  Each order line takes one path: stock or project.
  Rule 3  Material cost reaches a project only when a manufacturing order
          uses the material. So a line for a stock part must NOT carry the
          project tag (Odoo calls the tag "analytic distribution").

Odoo fills the project tag on every line by itself when the order has a
Project in its header. Test T3b in the test database showed the result:
$1,000 of glass became $2,000 of project cost (bill + manufacturing order).
`line_ids_to_clear` finds the lines where the tag must be removed.
"""

# Products that the PO Generator used for every line before Nov 2026.
# They stay as a fallback until the real products are in Odoo.
GENERIC_PRODUCT_NAMES = {
    "Glass": "SWR Glass Panel",
    "Aluminium": "SWR Aluminium Extrusion",
}

# Unit names that mean "this product is counted by area".
AREA_UNIT_NAMES = {"ft²", "ft2", "sq ft", "sqft", "sf", "m²", "m2"}


def is_area_unit(uom_name):
    """True when the Odoo unit of measure is an area (for example ft²)."""
    return str(uom_name or "").strip().lower() in AREA_UNIT_NAMES


def is_generic(product, po_type):
    """True when `product` is the old generic product for this PO type."""
    generic = GENERIC_PRODUCT_NAMES.get(po_type, "").lower()
    return bool(generic) and generic in str(product.get("name", "")).lower()


def match_candidates(products, code):
    """Odoo products whose code fits `code` (a profile number).

    Levels, best first:
      1. the code is the same,
      2. the Odoo code starts with the profile number (03111 -> 03111-MF),
      3. the profile number is inside the Odoo code.
    Returns all products of the best level that has a hit, or [].

    One profile often has several finishes in Odoo (03111-BA, 03111-CA,
    03111-MF ...). Then this returns more than one product, and a person
    must pick the finish. The app never guesses the finish.
    """
    want = str(code or "").strip().lower()
    if not want:
        return []
    coded = [p for p in products if str(p.get("default_code") or "").strip()]
    for test in (
        lambda c: c == want,
        lambda c: c.startswith(want),
        lambda c: want in c,
    ):
        hits = [p for p in coded if test(str(p["default_code"]).strip().lower())]
        if hits:
            return hits
    return []


def match_product(products, code):
    """The one product that fits `code`, or None when 0 or 2+ products fit."""
    hits = match_candidates(products, code)
    return hits[0] if len(hits) == 1 else None


def line_qty_and_price(po_type, line, product, price_per_unit):
    """Return (quantity, unit price) for one Odoo order line.

    Glass
      Product counted by area (ft²): quantity = total area, price = $/ft².
      Other product (for example the generic one, counted in pieces):
      quantity = pieces, price = line total / pieces. Same total both ways.

    Aluminium
      quantity = number of stock pieces, price = the unit price of the line.
      (Before Nov 2026 every aluminium line went to Odoo at $1.00.)
    """
    qty = float(line.get("qty") or 0)
    if po_type == "Glass":
        area_total = float(line.get("area_total") or 0)
        if is_area_unit(product.get("uom_name")):
            return area_total, float(price_per_unit or 0)
        line_total = area_total * float(price_per_unit or 0)
        return qty, (line_total / qty if qty > 0 else 0.0)
    return qty, float(line.get("unit_price") or 0)


def line_description(po_type, line):
    """Text for the Odoo order line. Keeps the tag, size and quantity."""
    if po_type == "Glass":
        parts = [
            str(line.get("description") or "").strip(),
            " ".join(x for x in (str(line.get("tag") or "").strip(),
                                 str(line.get("size_str") or "").strip()) if x),
            f"Area: {float(line.get('area_total') or 0):.3f} ft²  |  Pieces: {line.get('qty')}",
        ]
    else:
        parts = [
            f"{line.get('profile', '')}: {line.get('length_in', '')}\" extrusion".strip(),
            " ".join(x for x in (str(line.get("alloy") or "").strip(),
                                 str(line.get("finish") or "").strip()) if x),
            f"Pieces: {line.get('qty')}",
        ]
    return "\n".join(p for p in parts if p)


def build_order_lines(po_type, po_lines, price_per_unit, products_by_id, fallback_product_id):
    """Build the `order_line` value for purchase.order create.

    Each line may carry "product_id" (the Odoo product the person picked).
    A line with no product uses `fallback_product_id` (the generic product).
    Lines with quantity 0 are left out: Odoo would make an empty line.
    """
    out = []
    for line in po_lines:
        pid = line.get("product_id") or fallback_product_id
        product = products_by_id.get(pid, {"id": pid})
        qty, price = line_qty_and_price(po_type, line, product, price_per_unit)
        if qty <= 0:
            continue
        out.append((0, 0, {
            "product_id": pid,
            "name": line_description(po_type, line),
            "product_qty": qty,
            "price_unit": round(price, 6),
        }))
    return out


def order_total(order_lines):
    """Sum of quantity x price over lines made by `build_order_lines`."""
    return sum(v["product_qty"] * v["price_unit"] for _, _, v in order_lines)


def line_ids_to_clear(odoo_lines, storable_by_product):
    """Rule 3: ids of order lines whose project tag must be removed.

    odoo_lines            rows read from purchase.order.line with
                          "id", "product_id" ([id, name] or False) and
                          "analytic_distribution" (dict or False).
    storable_by_product   {product id: True when the product is kept in stock}

    A stock part with a tag is cleared. A service (freight, powder coating)
    keeps its tag: its cost goes to the project when the bill posts.
    """
    ids = []
    for row in odoo_lines:
        if not row.get("analytic_distribution"):
            continue
        prod = row.get("product_id")
        pid = prod[0] if isinstance(prod, (list, tuple)) and prod else prod
        if pid and storable_by_product.get(pid):
            ids.append(row["id"])
    return ids
