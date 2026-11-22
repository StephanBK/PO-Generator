"""Run the real Streamlit app against the fake Odoo, click through both flows,
and check what the app sends to Odoo."""
import pytest
from streamlit.testing.v1 import AppTest

import odoo_json2
from tests.fake_odoo import API_KEY, DB, FakeOdoo, PROJECT_ID


def _install(monkeypatch, fake):
    monkeypatch.setenv("ODOO_URL", "https://odoo.test")
    monkeypatch.setenv("ODOO_DB", DB)
    monkeypatch.setenv("ODOO_API_KEY", API_KEY)
    monkeypatch.setattr(odoo_json2.requests, "post", fake.post)
    import streamlit as st
    st.cache_data.clear()
    return fake


@pytest.fixture
def odoo(monkeypatch):
    return _install(monkeypatch, FakeOdoo())


@pytest.fixture
def odoo_no_generic(monkeypatch):
    """Odoo as it is today: the product "SWR Aluminium Extrusion" does not exist."""
    return _install(monkeypatch, FakeOdoo(hide_product_ids={910}))


def _select(at, label_start, value=None, index=None):
    box = next(b for b in at.selectbox if str(b.label).startswith(label_start))
    if value is None:
        value = box.options[index]
    box.select(value)
    return box


def _button(at, text):
    return next(b for b in at.button if text in str(b.label))


def test_glass_order_links_project_and_clears_stock_tags(odoo):
    at = AppTest.from_file("app.py", default_timeout=60).run()
    assert not at.exception
    _select(at, "Odoo Project", "INO-035 Test Job"); at.run()
    _select(at, "Select Glass file", index=0); at.run()
    assert not at.exception
    # one product dropdown per glass type; default is the generic product
    box1 = next(b for b in at.selectbox if str(b.label) == "Odoo product for GL-1")
    assert "generic" in box1.value
    box1.select(next(o for o in box1.options if "INOV-900" in o))       # real product, ft²
    box2 = next(b for b in at.selectbox if str(b.label) == "Odoo product for GL-2")
    box2.select(next(o for o in box2.options if "INOV-901" in o))       # real product, pieces
    at.run()
    _select(at, "Select Vendor", "Test Vendor")
    next(n for n in at.number_input if str(n.label).startswith("Price per ft")).set_value(10.0)
    at.run()
    _button(at, "Generate PO (.docx)").click(); at.run()
    assert not at.exception
    _button(at, "Create Draft PO in Odoo").click(); at.run()
    assert not at.exception
    assert not at.error, [e.value for e in at.error]

    vals = odoo.created_orders[0]
    assert vals["project_id"] == PROJECT_ID
    lines = [l[2] for l in vals["order_line"]]
    assert [l["product_id"] for l in lines] == [901, 902]
    assert (lines[0]["product_qty"], lines[0]["price_unit"]) == (40.0, 10.0)      # area product
    assert lines[1]["product_qty"] == 2 and round(lines[1]["price_unit"], 2) == 125.00  # 25 ft² x $10 / 2 pcs
    total = sum(l["product_qty"] * l["price_unit"] for l in lines)
    assert round(total, 2) == 650.00                                              # (40 + 25) x $10
    # rule 3: both lines are stock parts, both tags removed
    assert odoo.line_writes == [([700, 701], {"analytic_distribution": False})]
    assert "removed from 2 stock line" in odoo.chatter


def test_glass_order_with_generic_product_still_works_and_warns(odoo):
    at = AppTest.from_file("app.py", default_timeout=60).run()
    _select(at, "Odoo Project", "INO-035 Test Job"); at.run()
    _select(at, "Select Glass file", index=0); at.run()
    _select(at, "Select Vendor", "Test Vendor")
    next(n for n in at.number_input if str(n.label).startswith("Price per ft")).set_value(10.0)
    at.run()
    _button(at, "Generate PO (.docx)").click(); at.run()
    _button(at, "Create Draft PO in Odoo").click(); at.run()
    assert not at.exception and not at.error
    vals = odoo.created_orders[0]
    assert [l[2]["product_id"] for l in vals["order_line"]] == [900, 900]
    # the generic product is not a stock product: nothing to clear
    assert odoo.line_writes == []
    assert any("generic product" in str(w.value) for w in at.warning)


def test_aluminium_order_matches_products_and_sends_real_prices(odoo):
    at = AppTest.from_file("app.py", default_timeout=60).run()
    at.radio[0].set_value("Aluminium PO"); at.run()
    _select(at, "Odoo Project", "INO-035 Test Job"); at.run()
    _select(at, "Select cutting_list file", index=0); at.run()
    assert not at.exception
    boxes = [b for b in at.selectbox if str(b.key or "").startswith("al_prod_")]
    assert len(boxes) == 3
    assert boxes[0].value is None                      # 03111 has 2 finishes: the app does not guess
    assert "03111" in boxes[0].options[0] and "03111" in boxes[0].options[1]   # the 2 that fit come first
    boxes[0].select(next(o for o in boxes[0].options if "03111-MF" in o))
    assert "03210-CA" in boxes[1].value                # 1 fit: selected for the person
    assert "generic" in boxes[2].value                 # 99999 has no fit
    prices = [n for n in at.number_input if str(n.key or "").startswith("al_price_")]
    prices[0].set_value(40.0); prices[1].set_value(55.0); prices[2].set_value(5.0)
    _select(at, "Select Vendor", "Test Vendor")
    at.run()
    _button(at, "Generate Aluminium PO (.docx)").click(); at.run()
    _button(at, "Create Draft PO in Odoo").click(); at.run()
    assert not at.exception
    assert not at.error, [e.value for e in at.error]
    vals = odoo.created_orders[0]
    lines = [l[2] for l in vals["order_line"]]
    assert [(l["product_id"], l["product_qty"], l["price_unit"]) for l in lines] == \
        [(911, 3.0, 40.0), (912, 1.0, 55.0), (910, 1.0, 5.0)]
    # 3 x 40 + 1 x 55 + 1 x 5 = 180
    assert sum(l["product_qty"] * l["price_unit"] for l in lines) == 180.0
    # two stock sticks cleared; the generic line is not a stock product
    assert odoo.line_writes == [([700, 701], {"analytic_distribution": False})]




def test_no_product_and_no_generic_stops_with_a_clear_error(odoo_no_generic):
    odoo = odoo_no_generic
    at = AppTest.from_file("app.py", default_timeout=60).run()
    at.radio[0].set_value("Aluminium PO"); at.run()
    _select(at, "Odoo Project", "INO-035 Test Job"); at.run()
    _select(at, "Select cutting_list file", index=0); at.run()
    boxes = [b for b in at.selectbox if str(b.key or "").startswith("al_prod_")]
    assert boxes[2].value is None                      # 99999: no fit and no generic product in Odoo
    assert any("No Odoo product selected" in str(w.value) for w in at.warning)
    _select(at, "Select Vendor", "Test Vendor"); at.run()
    _button(at, "Generate Aluminium PO (.docx)").click(); at.run()
    _button(at, "Create Draft PO in Odoo").click(); at.run()
    assert not at.exception
    assert any("no Odoo product" in str(e.value) for e in at.error)
    assert odoo.created_orders == []                   # nothing was written to Odoo


def test_every_call_goes_through_json2_with_the_api_key(odoo):
    at = AppTest.from_file("app.py", default_timeout=60).run()
    assert not at.exception
    _select(at, "Odoo Project", "INO-035 Test Job"); at.run()
    assert odoo.requests, "the app made no Odoo call"
    assert all(r["url"].startswith("https://odoo.test/json/2/") for r in odoo.requests)
    assert all(r["headers"] == {"Authorization": f"bearer {API_KEY}", "X-Odoo-Database": DB} for r in odoo.requests)
    assert not any("xmlrpc" in r["url"] for r in odoo.requests)


def test_a_wrong_api_key_shows_an_error_and_makes_no_order(monkeypatch):
    fake = _install(monkeypatch, FakeOdoo())
    monkeypatch.setenv("ODOO_API_KEY", "wrong-key")
    at = AppTest.from_file("app.py", default_timeout=60).run()
    assert not at.exception
    shown = " ".join(str(e.value) for e in list(at.error) + list(at.warning))
    assert "refused the API key (401)" in shown
    assert fake.created_orders == []
