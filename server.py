"""
Caterpillar Asset Management V2 — MCP Server
============================================

Exposes the Cat Digital Asset Management V2 API (asset ownership records and
ownership transfer requests) as Model Context Protocol tools.

API:        https://services.cat.com/catDigital/assetManagement/v2
Auth:       OAuth2 client-credentials (Microsoft Entra ID)
Transport:  stdio (default, for VS Code / Inspector) or streamable-http (M365 Copilot)

Run:
    python server.py            # stdio
    python server.py --http     # streamable-http on :8765

All four operations are POST endpoints. Per the OpenAPI contract every write
operation requires a partyNumber plus an asset identifier (serialNumber +
makeCode/dealerMakeCode + dcn). The tools enforce the "one of makeCode or
dealerMakeCode" rule before the call leaves the process.
"""

from __future__ import annotations

import os
import sys
import time
import uuid
import threading
from typing import Any, Optional

import httpx
from dotenv import load_dotenv
from mcp.server.fastmcp import FastMCP

load_dotenv()

# --------------------------------------------------------------------------- #
# Configuration                                                               #
# --------------------------------------------------------------------------- #

BASE_URL = (
    os.environ.get("CAT_BASE_URL")
    or "https://services.cat.com/catDigital/assetManagement/v2"
).rstrip("/")

# Entra ID tenant baked into the OpenAPI spec — overridable via env.
# Use `or default` (not get's default arg) so an empty .env value falls back too.
TENANT_ID = os.environ.get("CAT_TENANT_ID") or "ceb177bf-013b-49ab-8a9c-4abce32afc1e"
# Caterpillar may supply a custom token URL; it takes precedence when set.
TOKEN_URL = (
    os.environ.get("CAT_TOKEN_URL")
    or f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
)

CLIENT_ID = os.environ.get("CAT_CLIENT_ID")
CLIENT_SECRET = os.environ.get("CAT_CLIENT_SECRET")
# Scope is "<resource-client-id>/.default" for client-credentials against the API.
SCOPE = os.environ.get("CAT_SCOPE", "")

# Some Cat Digital gateways additionally require an API subscription key header.
# Optional — only sent when set.
SUBSCRIPTION_KEY = os.environ.get("CAT_SUBSCRIPTION_KEY")
SUBSCRIPTION_HEADER = os.environ.get(
    "CAT_SUBSCRIPTION_HEADER", "x-api-key"
)

HTTP_TIMEOUT = float(os.environ.get("CAT_HTTP_TIMEOUT", "30"))

# Dealer code (partyNumber) resolution. Two optional env knobs:
#   CAT_FORCE_PARTY_NUMBER   — if set, OVERRIDES whatever the caller passes. Use
#                              during the testing phase to pin every call to your
#                              entitled test dealer code, even if the agent insists
#                              on sending a different (not-yet-entitled) code.
#   CAT_DEFAULT_PARTY_NUMBER — used only when the caller omits party_number.
# Once Caterpillar upgrades your credentials to your production dealer code,
# clear CAT_FORCE_PARTY_NUMBER so callers can specify any code they're entitled to.
FORCE_PARTY_NUMBER = os.environ.get("CAT_FORCE_PARTY_NUMBER") or ""
DEFAULT_PARTY_NUMBER = os.environ.get("CAT_DEFAULT_PARTY_NUMBER") or ""

mcp = FastMCP("cat-asset-management")


def _resolve_party(party_number: str) -> str:
    """Resolve the effective partyNumber: forced override > explicit > default."""
    if FORCE_PARTY_NUMBER:
        return FORCE_PARTY_NUMBER
    resolved = party_number or DEFAULT_PARTY_NUMBER
    if not resolved:
        raise ValueError(
            "party_number is required (no CAT_FORCE/DEFAULT_PARTY_NUMBER configured). "
            "Provide the dealer code, e.g. your test dealer code during onboarding."
        )
    return resolved


# --------------------------------------------------------------------------- #
# OAuth2 token management (client credentials, cached with refresh)           #
# --------------------------------------------------------------------------- #

class _TokenCache:
    """Thread-safe client-credentials token cache.

    Unlike the Snowflake PAT, Entra access tokens expire (~1h). We cache the
    token and refresh it ~60s before expiry.
    """

    def __init__(self) -> None:
        self._token: Optional[str] = None
        self._expires_at: float = 0.0
        self._lock = threading.Lock()

    def get(self) -> str:
        with self._lock:
            if self._token and time.time() < self._expires_at:
                return self._token
            return self._refresh()

    def _refresh(self) -> str:
        if not CLIENT_ID or not CLIENT_SECRET:
            raise RuntimeError(
                "Missing credentials. Set CAT_CLIENT_ID and CAT_CLIENT_SECRET "
                "in the environment or .env file."
            )
        if not SCOPE:
            raise RuntimeError(
                "Missing CAT_SCOPE. Set it to the API's '<client-ID>/.default' "
                "scope value from your Cat Digital onboarding."
            )

        resp = httpx.post(
            TOKEN_URL,
            data={
                "grant_type": "client_credentials",
                "client_id": CLIENT_ID,
                "client_secret": CLIENT_SECRET,
                "scope": SCOPE,
            },
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            timeout=HTTP_TIMEOUT,
        )
        if resp.status_code != 200:
            raise RuntimeError(
                f"Token request failed ({resp.status_code}): {resp.text[:500]}"
            )
        payload = resp.json()
        token = payload["access_token"]
        expires_in = int(payload.get("expires_in", 3600))
        self._token = token
        self._expires_at = time.time() + max(expires_in - 60, 30)
        return token


_token_cache = _TokenCache()


# --------------------------------------------------------------------------- #
# HTTP helper                                                                 #
# --------------------------------------------------------------------------- #

def _request(
    path: str,
    *,
    params: dict[str, Any],
    json_body: Optional[dict[str, Any]] = None,
    extra_headers: Optional[dict[str, str]] = None,
    tracking_id: Optional[str] = None,
) -> dict[str, Any]:
    """Perform an authenticated POST and return a normalized result dict.

    Returns a dict that always contains a ``tracking_id`` and either the parsed
    body (under ``data``) or a structured ``error``. Never raises on HTTP error
    status — surfaces the Cat Error schema to the agent instead.
    """
    # Drop unset query params so we don't send empty strings.
    clean_params = {k: v for k, v in params.items() if v not in (None, "")}

    tracking_id = tracking_id or str(uuid.uuid4())
    headers = {
        "Authorization": f"Bearer {_token_cache.get()}",
        "Accept": "application/json",
        "X-Cat-API-Tracking-Id": tracking_id,
    }
    if json_body is not None:
        headers["Content-Type"] = "application/json"
    if SUBSCRIPTION_KEY:
        headers[SUBSCRIPTION_HEADER] = SUBSCRIPTION_KEY
    if extra_headers:
        headers.update(extra_headers)

    url = f"{BASE_URL}{path}"
    try:
        resp = httpx.post(
            url, params=clean_params, json=json_body, headers=headers,
            timeout=HTTP_TIMEOUT,
        )
    except httpx.RequestError as exc:
        return {
            "ok": False,
            "tracking_id": tracking_id,
            "error": {"code": "network_error", "description": str(exc)},
        }

    resp_tracking = resp.headers.get("X-Cat-API-Tracking-Id", tracking_id)

    # 204 No Content — success with no body.
    if resp.status_code == 204:
        return {"ok": True, "status_code": 204, "tracking_id": resp_tracking, "data": None}

    # Parse body if any.
    body: Any = None
    if resp.content:
        try:
            body = resp.json()
        except ValueError:
            body = resp.text

    if 200 <= resp.status_code < 300:
        return {
            "ok": True,
            "status_code": resp.status_code,
            "tracking_id": resp_tracking,
            "data": body,
        }

    # Error — Cat APIs return the Error schema {code, description, details}.
    error = body if isinstance(body, dict) else {"description": body}
    return {
        "ok": False,
        "status_code": resp.status_code,
        "tracking_id": resp_tracking,
        "error": error,
    }


def _asset_params(
    party_number: str,
    serial_number: str,
    make_code: str,
    dealer_make_code: str,
    dcn: Optional[str] = None,
) -> dict[str, Any]:
    """Build & validate the common asset-identifier query params.

    Enforces the contract rule: exactly one of makeCode / dealerMakeCode.
    """
    if not make_code and not dealer_make_code:
        raise ValueError(
            "Provide make_code or dealer_make_code — the API requires one of them "
            "(error 400.202 if neither is supplied)."
        )
    params: dict[str, Any] = {
        "partyNumber": party_number,
        "serialNumber": serial_number,
        "makeCode": make_code,
        "dealerMakeCode": dealer_make_code,
    }
    if dcn is not None:
        params["dcn"] = dcn
    return params


# --------------------------------------------------------------------------- #
# Tools                                                                        #
# --------------------------------------------------------------------------- #

@mcp.tool()
def add_update_ownership_record(
    serial_number: str,
    dcn: str,
    party_number: str = "",
    make_code: str = "",
    dealer_make_code: str = "",
    ownership_type_code: str = "",
    model: str = "",
    model_year: str = "",
    product_family_code: str = "",
    product_family_name: str = "",
    base_asset_name: str = "",
    custom_asset_name: str = "",
    tracking_id: str = "",
) -> dict:
    """Add a new asset ownership record or update an existing one.

    Maps to POST /ownershipRecords. Provide exactly one of `make_code` or
    `dealer_make_code`.

    Behavior:
    * If the record exists, only the supplied body fields are updated.
    * If the record does NOT exist, `ownership_type_code`, `model`, and
      `model_year` become required.
    * Ownership conflict within a dealer auto-expires the conflicting record and
      this one becomes ACTIVE. Conflict across dealers creates a pending transfer
      request and this record is set to PENDING.

    Args:
        party_number: Organization with access to the equipment (query: partyNumber).
        serial_number: Asset serial number.
        dcn: Dealer Customer Number (dealer/customer association).
        make_code: Manufacturer code of the asset (e.g. "CW1"). One of make_code/dealer_make_code.
        dealer_make_code: 2-char dealer-specific make code (e.g. "CW"). One of make_code/dealer_make_code.
        ownership_type_code: owned|rental|leased|sold|inventory|unknown.
        model: Asset model (e.g. "980H").
        model_year: 4-digit year (e.g. "2006").
        product_family_code: Product family code (e.g. "MDWL").
        product_family_name: Product family name (e.g. "MEDIUM WHEEL LOADER").
        base_asset_name: Canonical asset name (dealer-settable).
        custom_asset_name: Custom asset name (takes priority over base name).
        tracking_id: Optional X-Cat-API-Tracking-Id; generated if omitted.

    Returns:
        dict with `ok`, `status_code`, `tracking_id`, and `data` (AddOwnershipResponse)
        or `error` (Cat Error schema).
    """
    params = _asset_params(_resolve_party(party_number), serial_number, make_code, dealer_make_code, dcn)

    body: dict[str, Any] = {}
    for key, value in (
        ("ownershipTypeCode", ownership_type_code),
        ("model", model),
        ("modelYear", model_year),
        ("productFamilyCode", product_family_code),
        ("productFamilyName", product_family_name),
        ("baseAssetName", base_asset_name),
        ("customAssetName", custom_asset_name),
    ):
        if value:
            body[key] = value

    return _request(
        "/ownershipRecords",
        params=params,
        json_body=body,
        tracking_id=tracking_id or None,
    )


@mcp.tool()
def expire_ownership_record(
    serial_number: str,
    dcn: str,
    party_number: str = "",
    make_code: str = "",
    dealer_make_code: str = "",
    tracking_id: str = "",
) -> dict:
    """Expire an existing asset ownership record.

    Maps to POST /ownershipRecords/expire. Provide exactly one of `make_code`
    or `dealer_make_code`. If a PENDING record is expired, its associated
    ownership transfer request is automatically cancelled. Returns 204 on success.

    Args:
        party_number: Organization with access to the equipment.
        serial_number: Asset serial number.
        dcn: Dealer Customer Number.
        make_code: Manufacturer code. One of make_code/dealer_make_code.
        dealer_make_code: Dealer-specific make code. One of make_code/dealer_make_code.
        tracking_id: Optional X-Cat-API-Tracking-Id.

    Returns:
        dict with `ok` and `tracking_id`; `data` is null on the 204 success path.
    """
    params = _asset_params(_resolve_party(party_number), serial_number, make_code, dealer_make_code, dcn)
    return _request(
        "/ownershipRecords/expire",
        params=params,
        json_body=None,
        tracking_id=tracking_id or None,
    )


@mcp.tool()
def search_ownership_records(
    dcn: str = "",
    serial_number: str = "",
    asset_name: str = "",
    make_code: str = "",
    party_number: str = "",
    case_sensitive: bool = False,
    sort_by: Optional[list[str]] = None,
    order_by: Optional[list[str]] = None,
    response_attributes: Optional[list[str]] = None,
    filters: Optional[list[dict]] = None,
    user_preferences: str = "",
    tracking_id: str = "",
) -> dict:
    """Search ACTIVE or PENDING asset ownership records.

    Maps to POST /ownershipRecords/search. Provide one or two of the search
    fields below — they are matched **exactly** (not partial/contains). When you
    give two, they combine with logical AND.

    You can search by ANY of these four fields, alone or in pairs:
      * dcn            — Dealer Customer Number, e.g. "12345"
      * serial_number  — asset serial number, e.g. "2WS23456"
      * asset_name     — asset name
      * make_code      — manufacturer code, e.g. "CW1"
                         (make_code can't be used alone — pair it with another field)

    Examples:
      * Search by DCN:            dcn="12345"
      * Search by serial:         serial_number="2WS23456"
      * Search by make + serial:  make_code="CW1", serial_number="2WS23456"

    Args:
        dcn: Filter by Dealer Customer Number (exact match).
        serial_number: Filter by asset serial number (exact match).
        asset_name: Filter by asset name (exact match).
        make_code: Filter by manufacturer code (exact match; needs a 2nd field).
        party_number: Dealer code; usually leave blank (the server supplies it).
        case_sensitive: Match values case-sensitively (default false).
        sort_by: Up to 3 of: dcn | assetName | serialNumber | makeCode.
        order_by: Up to 3 of ASC | DESC, positionally matched to sort_by.
        response_attributes: Optional dotted attribute paths to return
            (e.g. ["metadata", "ownership.dealerAssociation.dcn"]). Empty = all.
        filters: Advanced escape hatch — raw filter objects. Ignored unless the
            flat fields above are all empty. Most callers should not use this.
        user_preferences: Optional base64 X-Cat-User-Preferences header value.
        tracking_id: Optional X-Cat-API-Tracking-Id.

    Returns:
        dict with `ok`, `tracking_id`, and `data` (an object with
        `ownershipRecords`) or `error`.
    """
    # Build filters from the flat fields (the normal path).
    field_map = [
        ("dcn", dcn),
        ("serialNumber", serial_number),
        ("assetName", asset_name),
        ("makeCode", make_code),
    ]
    provided = [(prop, val) for prop, val in field_map if val]

    if provided:
        if len(provided) > 2:
            raise ValueError(
                "Search supports at most 2 fields at once (API limit). "
                "Choose up to two of: dcn, serial_number, asset_name, make_code."
            )
        names = [p for p, _ in provided]
        if "makeCode" in names and len(provided) < 2:
            raise ValueError(
                "make_code can't be the only search field — pair it with "
                "serial_number, dcn, or asset_name."
            )
        built_filters = [
            {
                "type": "stringEquals",
                "propertyName": prop,
                "values": [val],
                "isCaseSensitive": case_sensitive,
            }
            for prop, val in provided
        ]
    elif filters:
        # Advanced fallback: caller supplied raw filter objects directly.
        if len(filters) > 2:
            raise ValueError("A maximum of 2 filters is allowed.")
        raw_props = [str(f.get("propertyName", "")).lower() for f in filters]
        if "makecode" in raw_props and len(filters) < 2:
            raise ValueError(
                "When filtering by makeCode, at least one additional filter is required."
            )
        built_filters = filters
    else:
        raise ValueError(
            "Provide at least one search field: dcn, serial_number, asset_name, "
            "or make_code."
        )

    params: dict[str, Any] = {"partyNumber": _resolve_party(party_number)}
    if sort_by:
        params["sortBy"] = ",".join(sort_by)
    if order_by:
        params["orderBy"] = ",".join(order_by)

    body: dict[str, Any] = {"filters": built_filters}
    if response_attributes:
        body["responseAttributes"] = response_attributes

    extra = {"X-Cat-User-Preferences": user_preferences} if user_preferences else None

    return _request(
        "/ownershipRecords/search",
        params=params,
        json_body=body,
        extra_headers=extra,
        tracking_id=tracking_id or None,
    )


@mcp.tool()
def approve_deny_ownership_transfer(
    serial_number: str,
    status: str,
    party_number: str = "",
    make_code: str = "",
    dealer_make_code: str = "",
    reason: str = "",
    tracking_id: str = "",
) -> dict:
    """Approve or reject a pending asset ownership transfer request.

    Maps to POST /ownershipRequests/transfer. Only the dealer currently owning
    the asset can act. Provide exactly one of `make_code`/`dealer_make_code`.
    Returns 204 on success.

    * APPROVED: expires the owning dealer's OWNED ownership and creates a new
      OWNED record for the recommended dealer.
    * REJECTED: retains current ownership and expires the pending association.
      A `reason` MUST be provided when rejecting.

    Args:
        party_number: Organization with access to the equipment.
        serial_number: Asset serial number.
        status: "APPROVED" or "REJECTED".
        make_code: Manufacturer code. One of make_code/dealer_make_code.
        dealer_make_code: Dealer-specific make code. One of make_code/dealer_make_code.
        reason: Required when REJECTED; optional otherwise.
        tracking_id: Optional X-Cat-API-Tracking-Id.

    Returns:
        dict with `ok` and `tracking_id`; `data` is null on the 204 success path.
    """
    status = status.upper()
    if status not in ("APPROVED", "REJECTED"):
        raise ValueError('status must be "APPROVED" or "REJECTED".')
    if status == "REJECTED" and not reason:
        raise ValueError("A reason must be provided when rejecting a transfer.")

    params = _asset_params(_resolve_party(party_number), serial_number, make_code, dealer_make_code)

    body: dict[str, Any] = {"status": status}
    if reason:
        body["reason"] = reason

    return _request(
        "/ownershipRequests/transfer",
        params=params,
        json_body=body,
        tracking_id=tracking_id or None,
    )


# --------------------------------------------------------------------------- #
# Entry point                                                                 #
# --------------------------------------------------------------------------- #

if __name__ == "__main__":
    if "--http" in sys.argv:
        port = int(os.environ.get("CAT_MCP_PORT", "8765"))
        mcp.settings.host = "0.0.0.0"
        mcp.settings.port = port
        mcp.run(transport="streamable-http")
    else:
        mcp.run()
