"""The JSON-2 connection: argument names, headers, and what happens when Odoo refuses."""
import pytest
import requests

import odoo_json2
from odoo_json2 import OdooError, to_body


def test_positional_arguments_get_their_names():
    assert to_body("search_read", [[("active", "=", True)]], {"fields": ["id", "name"], "limit": 200}) == \
        {"domain": [("active", "=", True)], "fields": ["id", "name"], "limit": 200}
    assert to_body("search", [[("id", "in", [1, 2])]], {"limit": 1}) == {"domain": [("id", "in", [1, 2])], "limit": 1}
    assert to_body("read", [[77]], {"fields": ["name"]}) == {"ids": [77], "fields": ["name"]}
    assert to_body("write", [[701, 702], {"analytic_distribution": False}]) == \
        {"ids": [701, 702], "vals": {"analytic_distribution": False}}
    assert to_body("message_post", [[77]], {"body": "hello"}) == {"ids": [77], "body": "hello"}


def test_create_always_sends_a_list_of_records():
    assert to_body("create", [{"partner_id": 300}]) == {"vals_list": [{"partner_id": 300}]}
    assert to_body("create", [[{"a": 1}, {"a": 2}]]) == {"vals_list": [{"a": 1}, {"a": 2}]}


def test_a_method_the_app_does_not_use_is_refused():
    # unlink (delete) is not in the list: the app cannot delete an Odoo record, even by a coding mistake
    with pytest.raises(OdooError, match="unlink"):
        to_body("unlink", [[77]])
    with pytest.raises(OdooError, match="2 positional"):
        to_body("read", [[77], ["name"]])
    with pytest.raises(OdooError, match="two times"):
        to_body("read", [[77]], {"ids": [78]})


class _Resp:
    def __init__(self, status, data, text=""):
        self.status_code, self._data, self.text = status, data, text or str(data)

    def json(self):
        if isinstance(self._data, Exception):
            raise self._data
        return self._data


def test_call_sends_the_key_and_the_database_in_the_headers(monkeypatch):
    seen = {}

    def post(url, json=None, headers=None, timeout=None):
        seen.update(url=url, json=json, headers=headers, timeout=timeout)
        return _Resp(200, [{"id": 12, "name": "INO-035"}])
    monkeypatch.setattr(odoo_json2.requests, "post", post)
    out = odoo_json2.call("https://x.odoo.com/", "db1", "k123", "project.project", "search_read",
                          [[("active", "=", True)]], {"fields": ["name"]})
    assert out == [{"id": 12, "name": "INO-035"}]
    assert seen["url"] == "https://x.odoo.com/json/2/project.project/search_read"
    assert seen["headers"] == {"Authorization": "bearer k123", "X-Odoo-Database": "db1"}
    assert seen["json"] == {"domain": [("active", "=", True)], "fields": ["name"]}


def test_errors_are_plain(monkeypatch):
    with pytest.raises(OdooError, match="not configured"):
        odoo_json2.call("https://x.odoo.com", "db1", "", "res.partner", "search_read", [[]])

    monkeypatch.setattr(odoo_json2.requests, "post", lambda *a, **k: _Resp(401, {"message": "Invalid apikey"}))
    with pytest.raises(OdooError, match=r"refused the API key \(401\)"):
        odoo_json2.call("https://x.odoo.com", "db1", "bad", "res.partner", "search_read", [[]])

    monkeypatch.setattr(odoo_json2.requests, "post",
                        lambda *a, **k: _Resp(422, {"message": "Invalid field 'nope' on model 'res.partner'"}))
    with pytest.raises(OdooError, match="returned 422: Invalid field 'nope'"):
        odoo_json2.call("https://x.odoo.com", "db1", "k", "res.partner", "search_read", [[]], {"fields": ["nope"]})

    monkeypatch.setattr(odoo_json2.requests, "post", lambda *a, **k: _Resp(502, ValueError("no json"), "Bad Gateway"))
    with pytest.raises(OdooError, match="returned 502: Bad Gateway"):
        odoo_json2.call("https://x.odoo.com", "db1", "k", "res.partner", "search_read", [[]])

    def down(*a, **k):
        raise requests.ConnectionError("dns")
    monkeypatch.setattr(odoo_json2.requests, "post", down)
    with pytest.raises(OdooError, match="did not answer"):
        odoo_json2.call("https://x.odoo.com", "db1", "k", "res.partner", "search_read", [[]])
