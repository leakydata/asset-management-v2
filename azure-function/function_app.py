"""
Cat Asset Management V2 - read-only broker (Azure Function).

Exposes GET/POST /api/search?serial=...&dcn=... and calls the Cat search
endpoint server-side, so callers never see the Cat OAuth credentials. Only
read (search) is exposed - there are no write routes.

Credentials come from environment / app settings (Key Vault references in
Azure, local.settings.json locally):
  CAT_CLIENT_ID, CAT_CLIENT_SECRET, CAT_SCOPE,
  CAT_TENANT_ID (or CAT_TOKEN_URL), CAT_DEFAULT_PARTY_NUMBER.

Auth level is FUNCTION (a function key is required to call it). Add Entra
"Easy Auth" in the portal for per-user authentication.
"""
import os
import json
import time
import threading

import azure.functions as func
import httpx

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

BASE_URL = (os.environ.get("CAT_BASE_URL")
            or "https://services.cat.com/catDigital/assetManagement/v2").rstrip("/")
TENANT_ID = os.environ.get("CAT_TENANT_ID") or "ceb177bf-013b-49ab-8a9c-4abce32afc1e"
TOKEN_URL = (os.environ.get("CAT_TOKEN_URL")
             or f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token")
SCOPE = os.environ.get("CAT_SCOPE", "")
PARTY = (os.environ.get("CAT_FORCE_PARTY_NUMBER")
         or os.environ.get("CAT_DEFAULT_PARTY_NUMBER") or "")

_token = None
_token_exp = 0.0
_lock = threading.Lock()


def _get_token() -> str:
    """Client-credentials token, cached in-process and refreshed ~60s early."""
    global _token, _token_exp
    with _lock:
        if _token and time.time() < _token_exp:
            return _token
        cid = os.environ.get("CAT_CLIENT_ID")
        sec = os.environ.get("CAT_CLIENT_SECRET")
        if not cid or not sec or not SCOPE:
            raise RuntimeError("Missing CAT_CLIENT_ID / CAT_CLIENT_SECRET / CAT_SCOPE.")
        r = httpx.post(
            TOKEN_URL,
            data={"grant_type": "client_credentials",
                  "client_id": cid, "client_secret": sec, "scope": SCOPE},
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            timeout=30,
        )
        r.raise_for_status()
        p = r.json()
        _token = p["access_token"]
        _token_exp = time.time() + int(p.get("expires_in", 3600)) - 60
        return _token


@app.route(route="search", methods=["GET", "POST"])
def search(req: func.HttpRequest) -> func.HttpResponse:
    """Search ownership records by serial and/or DCN (exact match)."""
    try:
        body = req.get_json()
    except ValueError:
        body = {}
    serial = (req.params.get("serial") or (body or {}).get("serial") or "").strip()
    dcn = (req.params.get("dcn") or (body or {}).get("dcn") or "").strip()

    if not serial and not dcn:
        return func.HttpResponse(
            json.dumps({"error": "Provide a serial and/or dcn query parameter."}),
            status_code=400, mimetype="application/json")

    filters = []
    if dcn:
        filters.append({"type": "stringEquals", "propertyName": "dcn", "values": [dcn]})
    if serial:
        filters.append({"type": "stringEquals", "propertyName": "serialNumber", "values": [serial]})

    try:
        token = _get_token()
    except Exception as exc:
        return func.HttpResponse(
            json.dumps({"error": f"auth failed: {exc}"}),
            status_code=502, mimetype="application/json")

    try:
        resp = httpx.post(
            f"{BASE_URL}/ownershipRecords/search",
            params={"partyNumber": PARTY},
            json={"filters": filters},
            headers={"Authorization": f"Bearer {token}",
                     "Accept": "application/json",
                     "Content-Type": "application/json"},
            timeout=30,
        )
    except httpx.RequestError as exc:
        return func.HttpResponse(
            json.dumps({"error": f"upstream request failed: {exc}"}),
            status_code=502, mimetype="application/json")

    # Pass the Cat response straight through (status + JSON body).
    return func.HttpResponse(
        resp.text or "{}", status_code=resp.status_code, mimetype="application/json")
