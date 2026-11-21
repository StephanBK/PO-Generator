import odoo_rules as rules

GLASS_AREA = {"id": 901, "name": "GLASS VIG Example", "uom_name": "ft²"}
GLASS_PCS = {"id": 902, "name": "GLASS MONO Example", "uom_name": "Units"}
GENERIC = {"id": 900, "name": "SWR Glass Panel", "uom_name": "ft²"}


def test_area_unit_names():
    assert rules.is_area_unit("ft²") and rules.is_area_unit(" FT2 ")
    assert not rules.is_area_unit("Units") and not rules.is_area_unit(None)


def test_glass_area_product_uses_area_and_price_per_sqft():
    line = {"qty": 5, "area_total": 40.0}
    assert rules.line_qty_and_price("Glass", line, GLASS_AREA, 9.5) == (40.0, 9.5)


def test_glass_piece_product_keeps_the_same_total():
    line = {"qty": 5, "area_total": 40.0}
    qty, price = rules.line_qty_and_price("Glass", line, GLASS_PCS, 9.5)
    assert qty == 5
    assert round(qty * price, 6) == round(40.0 * 9.5, 6)   # 380.00 both ways


def test_aluminium_uses_the_line_unit_price_not_one_dollar():
    line = {"qty": 3, "unit_price": 41.25, "profile": "03111", "length_in": 196}
    assert rules.line_qty_and_price("Aluminium", line, {"uom_name": "Units"}, None) == (3.0, 41.25)


def test_build_order_lines_total_and_fallback():
    lines = [{"qty": 5, "area_total": 40.0, "tag": "W1", "size_str": '24" x 48"', "product_id": 901},
             {"qty": 2, "area_total": 25.0, "tag": "W2", "size_str": '30" x 60"'},      # no product: fallback
             {"qty": 0, "area_total": 0.0, "tag": "W3"}]                                 # dropped
    out = rules.build_order_lines("Glass", lines, 10.0, {901: GLASS_AREA, 900: GENERIC}, 900)
    assert [v["product_id"] for _, _, v in out] == [901, 900]
    assert round(rules.order_total(out), 2) == 650.00       # (40 + 25) x $10
    assert "W1" in out[0][2]["name"]


def test_match_prefers_exact_then_prefix_then_inside():
    products = [{"id": 1, "default_code": "03111-MF"}, {"id": 2, "default_code": "03111"},
                {"id": 3, "default_code": "X-03210-CA"}, {"id": 4, "default_code": ""}]
    assert rules.match_product(products, "03111")["id"] == 2            # exact wins
    assert rules.match_product(products[:1], "03111")["id"] == 1        # prefix
    assert rules.match_product(products, "03210")["id"] == 3            # inside
    assert rules.match_product(products, "99999") is None
    assert rules.match_product(products, "") is None


def test_several_finishes_are_never_guessed():
    products = [{"id": 1, "default_code": "03111-BA"}, {"id": 2, "default_code": "03111-MF"}]
    assert [p["id"] for p in rules.match_candidates(products, "03111")] == [1, 2]
    assert rules.match_product(products, "03111") is None


def test_rule_3_clears_stock_lines_and_keeps_service_lines():
    odoo_lines = [
        {"id": 1, "product_id": [901, "glass"], "analytic_distribution": {"11": 100}},   # stock, tagged
        {"id": 2, "product_id": [920, "freight"], "analytic_distribution": {"11": 100}}, # service, tagged
        {"id": 3, "product_id": [911, "stick"], "analytic_distribution": False},         # stock, no tag
        {"id": 4, "product_id": False, "analytic_distribution": {"11": 100}},            # note line
    ]
    storable = {901: True, 920: False, 911: True}
    assert rules.line_ids_to_clear(odoo_lines, storable) == [1]


def test_cut_list_prefix_is_ignored_when_the_number_fits():
    products = [{"id": 1, "default_code": "03210-CA"}, {"id": 2, "default_code": "03111-MF"}]
    assert rules.match_product(products, "SWR-IG-03210")["id"] == 1
    assert rules.match_candidates(products, "SWR-IG-03003") == []      # no such number in Odoo
