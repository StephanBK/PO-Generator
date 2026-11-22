"""A small fake Odoo for tests. It answers the JSON-2 calls that app.py makes
(HTTPS POST to /json/2/<model>/<method>) and behaves like Odoo 19 in one important
way: when an order is created with a project, every line gets the project tag
(analytic distribution)."""
import base64
import io

import pandas as pd

PROJECT_ID = 12
API_KEY = "test-key"
DB = "inovues"
ANALYTIC_ID = 11

PRODUCTS = [
    # id, name, code, uom, storable, category
    (900, "SWR Glass Panel", "", "ft²", False, "All"),
    (901, "GLASS VIG Example 10mm", "INOV-900", "ft²", True, "All / INOVUES / Glass"),
    (902, "GLASS MONOLITHIC Example 6mm", "INOV-901", "Units", True, "All / INOVUES / Glass"),
    (910, "SWR Aluminium Extrusion", "", "Units", False, "All"),
    (911, "Aluminum Profile 03111 MF, 16 ft 4 in", "03111-MF", "Units", True, "All / INOVUES / Aluminum Profiles & Plates"),
    (912, "Aluminum Profile 03210 CA, 20 ft 6 in", "03210-CA", "Units", True, "All / INOVUES / Aluminum Profiles & Plates"),
    (913, "Aluminum Profile 03111 BA, 16 ft 4 in", "03111-BA", "Units", True, "All / INOVUES / Aluminum Profiles & Plates"),
    (920, "Freight", "", "Units", False, "All / Services"),
]


def glass_xlsx():
    rows = [[None] * 7 for _ in range(12)]
    rows.append(["Tag", "Glass Type", "Glass Width (1/16)", "Glass Height (1/16)",
                 "Area Each (ft²)", "Qty", "Area Total (ft²)"])
    rows.append(["W1", "GL-1", "24 1/16", "48 1/16", 8.0, 5, 40.0])
    rows.append(["W2", "GL-2", "30", "60", 12.5, 2, 25.0])
    rows.append(["Totals", None, None, None, None, 7, 65.0])
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame(rows).to_excel(w, sheet_name="Glass", header=False, index=False)
    return buf.getvalue()


def optimizer_xlsx():
    rows = [["SUMMARY", None, None], [None, None, None],
            ["DETAILED CUT LIST", None, None],
            ["Part Number", "Stock Length (ft)", "Cuts"],
            ["03111", 16.333, "a"], ["03111", 16.333, "b"], ["03111", 16.333, "c"],
            ["03210", 20.5, "d"], ["99999", 8, "e"]]
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame(rows).to_excel(w, sheet_name="Sheet1", header=False, index=False)
    return buf.getvalue()


class FakeOdoo:
    def __init__(self, hide_product_ids=()):
        self.hidden = set(hide_product_ids)
        self.calls = []
        self.created_orders = []
        self.line_writes = []
        self.lines = []
        self.requests = []

    def post(self, url, json=None, headers=None, timeout=None):
        """Stands in for requests.post. Checks the request like Odoo does, then answers."""
        self.requests.append({"url": url, "headers": dict(headers or {}), "body": json})
        if (headers or {}).get("Authorization") != f"bearer {API_KEY}" or (headers or {}).get("X-Odoo-Database") != DB:
            return _Response(401, {"message": "Invalid apikey"})
        model, method = url.split("/json/2/")[1].split("/")
        body = dict(json or {})
        names = {"search_read": ["domain"], "search": ["domain"], "read": ["ids"], "write": ["ids", "vals"],
                 "create": ["vals_list"], "message_post": ["ids"]}[method]
        args = [body.pop(n) for n in names if n in body]
        if method == "create":
            assert isinstance(args[0], list), "JSON-2 create takes vals_list, a list"
            return _Response(200, [self.execute(model, method, [v], body) for v in args[0]])
        return _Response(200, self.execute(model, method, args, body))

    def execute(self, model, method, args, kwargs):
        self.calls.append((model, method))
        h = getattr(self, f"_{model.replace('.', '_')}__{method}", None)
        if h is None:
            raise AssertionError(f"FakeOdoo has no answer for {model}.{method}")
        return h(args, kwargs)

    # project, task, attachment
    def _project_project__search_read(self, args, kw):
        return [{"id": PROJECT_ID, "name": "INO-035 Test Job"}]

    def _project_task__search_read(self, args, kw):
        return [{"id": 5, "name": "SWR Cutlist / Cutting Optimization", "create_date": "2026-11-01"}]

    def _ir_attachment__search_read(self, args, kw):
        return [
            {"id": 1, "name": "INO-035_SWR_Glass_test.xlsx", "res_id": 5, "create_date": "2026-11-01",
             "datas": base64.b64encode(glass_xlsx()).decode()},
            {"id": 2, "name": "cutting_list_test.xlsx", "res_id": 5, "create_date": "2026-11-01",
             "datas": base64.b64encode(optimizer_xlsx()).decode()},
        ]

    def _ir_attachment__create(self, args, kw):
        return 501

    # partners, terms, users
    def _res_partner__search_read(self, args, kw):
        return [{"id": 300, "name": "Test Vendor", "email": "v@example.com", "phone": "", "street": "1 Main St",
                 "street2": False, "city": "Houston", "state_id": [1, "Texas"], "zip": "77001",
                 "country_id": [233, "United States"], "child_ids": []}]

    def _account_payment_term__search_read(self, args, kw):
        return [{"id": 3, "name": "30 Days"}]

    def _account_incoterms__search_read(self, args, kw):
        return [{"id": 4, "code": "DAP", "name": "DELIVERED AT PLACE"}]

    def _res_users__search_read(self, args, kw):
        return [{"id": 2, "name": "Stephan Ketterer", "login": "sketterer@inovues.com"}]

    # products
    def _product_product__search_read(self, args, kw):
        domain = args[0]
        word = next(c[2] for c in domain if isinstance(c, (list, tuple)) and c[0] == "categ_id.complete_name")
        generic = next(c[2] for c in domain if isinstance(c, (list, tuple)) and c[0] == "name")
        out = []
        for pid, name, code, uom, storable, categ in PRODUCTS:
            if pid in self.hidden:
                continue
            if generic.lower() in name.lower() or word.lower() in categ.lower():
                out.append({"id": pid, "name": name, "default_code": code or False,
                            "uom_id": [1, uom], "is_storable": storable})
        return out

    def _product_product__search(self, args, kw):
        word = args[0][0][2].lower()
        return [p[0] for p in PRODUCTS if word in p[1].lower() and p[0] not in self.hidden]

    def _product_product__read(self, args, kw):
        ids = args[0]
        return [{"id": p[0], "is_storable": p[4]} for p in PRODUCTS if p[0] in ids]

    # purchase order
    def _purchase_order__create(self, args, kw):
        vals = args[0]
        self.created_orders.append(vals)
        tag = {str(ANALYTIC_ID): 100} if vals.get("project_id") else False
        self.lines = [{"id": 700 + i, "product_id": [line[2]["product_id"], "x"],
                       "analytic_distribution": tag, "vals": line[2]}
                      for i, line in enumerate(vals["order_line"])]
        return 77

    def _purchase_order__read(self, args, kw):
        return [{"id": 77, "name": "P00099"}]

    def _purchase_order__message_post(self, args, kw):
        self.chatter = kw["body"]
        return 1

    def _purchase_order_line__search_read(self, args, kw):
        return [{k: v for k, v in line.items() if k != "vals"} for line in self.lines]

    def _purchase_order_line__write(self, args, kw):
        ids, vals = args
        self.line_writes.append((list(ids), vals))
        for line in self.lines:
            if line["id"] in ids:
                line.update(vals)
        return True


class _Response:
    def __init__(self, status_code, data):
        self.status_code, self._data, self.text = status_code, data, str(data)

    def json(self):
        return self._data
