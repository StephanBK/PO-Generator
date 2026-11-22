"""Odoo connection through the JSON-2 API (Odoo 19 and later).

Why: Odoo removes the old XML-RPC interface in Odoo Online 21.1. JSON-2 is its replacement.

Every call is one HTTPS POST:
    POST {ODOO_URL}/json/2/{model}/{method}
    Authorization: bearer {ODOO_API_KEY}
    X-Odoo-Database: {ODOO_DB}
    body: the arguments of the method, by NAME, as JSON

There is no login step. The API key identifies the Odoo user on every call.

XML-RPC took the main arguments by POSITION: execute_kw(db, uid, key, model, method, args, kwargs).
JSON-2 takes them by name. `to_body` gives each positional argument its name, so the rest of
the app keeps its calls as they are:

    search_read([domain], {"fields": [...]})     -> {"domain": domain, "fields": [...]}
    read([ids], {"fields": [...]})               -> {"ids": ids, "fields": [...]}
    write([ids, vals])                           -> {"ids": ids, "vals": vals}
    create([vals])                               -> {"vals_list": [vals]}   (answer: a list of ids)
    message_post([ids], {"body": ...})           -> {"ids": ids, "body": ...}
"""
from __future__ import annotations

import requests

# The names of the positional arguments, per method. Only the methods the app uses.
POSITIONAL = {
    "search_read": ("domain",),
    "search": ("domain",),
    "search_count": ("domain",),
    "read": ("ids",),
    "write": ("ids", "vals"),
    "create": ("vals_list",),
    "message_post": ("ids",),
}


class OdooError(Exception):
    pass


def to_body(method: str, args=(), kwargs=None) -> dict:
    """The JSON body for one call, from XML-RPC style (args, kwargs)."""
    names = POSITIONAL.get(method)
    if names is None:
        raise OdooError(f"Odoo method '{method}' is not in the list of methods this app may call")
    args = list(args or [])
    if len(args) > len(names):
        raise OdooError(f"Odoo method '{method}' got {len(args)} positional arguments; it takes {len(names)}")
    body = dict(kwargs or {})
    for name, value in zip(names, args):
        if name in body:
            raise OdooError(f"Odoo method '{method}': argument '{name}' was given two times")
        body[name] = value
    if method == "create" and isinstance(body.get("vals_list"), dict):
        body["vals_list"] = [body["vals_list"]]          # one record: JSON-2 still wants a list
    return body


def call(url: str, db: str, key: str, model: str, method: str, args=(), kwargs=None, timeout: float = 120.0):
    """One call to Odoo. Raises OdooError with Odoo's own message when Odoo refuses."""
    if not (url and db and key):
        raise OdooError("Odoo is not configured: set ODOO_URL, ODOO_DB and ODOO_API_KEY.")
    body = to_body(method, args, kwargs)
    try:
        r = requests.post(f"{url.rstrip('/')}/json/2/{model}/{method}", json=body,
                          headers={"Authorization": f"bearer {key}", "X-Odoo-Database": db}, timeout=timeout)
    except requests.RequestException as e:
        raise OdooError(f"Odoo did not answer ({model}.{method}): {type(e).__name__}") from e
    if r.status_code >= 400:
        try:
            msg = r.json().get("message") or r.text
        except Exception:  # noqa: BLE001 - the error page was not JSON
            msg = r.text
        if r.status_code in (401, 403):
            raise OdooError(f"Odoo refused the API key ({r.status_code}) for {model}.{method}: {str(msg)[:200]}")
        raise OdooError(f"Odoo {model}.{method} returned {r.status_code}: {str(msg)[:300]}")
    return r.json()
