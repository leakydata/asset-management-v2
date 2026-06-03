# Cat Asset Management V2 — MCP Server

An [MCP](https://modelcontextprotocol.io) server wrapping the Caterpillar Digital
**Asset Management V2** API. It exposes the API's four operations as tools any MCP
client (VS Code, MCP Inspector, M365 Copilot Studio) can call.

| Tool | Endpoint | Purpose |
|------|----------|---------|
| `add_update_ownership_record` | `POST /ownershipRecords` | Add or update an asset ownership record |
| `expire_ownership_record` | `POST /ownershipRecords/expire` | Expire an ownership record |
| `search_ownership_records` | `POST /ownershipRecords/search` | Search ACTIVE/PENDING records |
| `approve_deny_ownership_transfer` | `POST /ownershipRequests/transfer` | Approve/reject a pending transfer |

## Setup

Managed with [uv](https://docs.astral.sh/uv/):

```bash
uv sync                    # creates .venv and installs dependencies
cp .env.example .env       # then fill in CAT_CLIENT_ID / CAT_CLIENT_SECRET / CAT_SCOPE
```

### Credentials

Auth is OAuth2 **client-credentials** against Microsoft Entra ID. From your Cat
Digital app onboarding you need:

- `CAT_CLIENT_ID` — your app registration's client ID
- `CAT_CLIENT_SECRET` — its client secret
- `CAT_SCOPE` — the **target API's** client ID + `/.default` (e.g. `<api-client-id>/.default`)

The tenant (`ceb177bf-…`) and base URL default from the API spec; override via
`.env` if needed. Some gateways also require a subscription key — set
`CAT_SUBSCRIPTION_KEY` (and `CAT_SUBSCRIPTION_HEADER` if it isn't `x-api-key`).

The server fetches and **caches** the access token, refreshing ~60s before expiry.

## Running

```bash
uv run server.py            # stdio — for VS Code & MCP Inspector
uv run server.py --http     # streamable-http on :8765 — for M365 Copilot Studio
```

### MCP Inspector

```bash
npx @modelcontextprotocol/inspector uv run server.py
```

### VS Code (`.vscode/mcp.json` or settings)

```json
{
  "servers": {
    "cat-asset-management": {
      "command": "uv",
      "args": ["run", "--directory", "${workspaceFolder}", "server.py"]
    }
  }
}
```

### M365 Copilot Studio

Run `uv run server.py --http`, expose `:8765` via a dev tunnel, and register
`https://<tunnel>/mcp` as an MCP action. (Copilot Studio requires streamable HTTP.)

## Tool contract notes

These mirror the OpenAPI rules and are enforced before the request is sent:

- **One of** `make_code` **or** `dealer_make_code` must be provided (never both).
- `search_ownership_records`: 1–2 filters; if filtering by `makeCode`, at least
  one other filter is required. Filter shape:
  ```json
  {"type": "stringEquals", "propertyName": "dcn", "values": ["12345"], "isCaseSensitive": true}
  ```
  `propertyName` ∈ `dcn | assetName | serialNumber | makeCode`.
- `approve_deny_ownership_transfer`: `status` is `APPROVED`/`REJECTED`; a `reason`
  is required when rejecting.
- New ownership records (when the record doesn't yet exist) require
  `ownership_type_code`, `model`, and `model_year`.

Every tool returns a normalized dict:

```jsonc
{ "ok": true, "status_code": 200, "tracking_id": "…", "data": { … } }   // success
{ "ok": false, "status_code": 400, "tracking_id": "…", "error": { "code": "400.001", "description": "…" } }  // error
```

204 responses (expire, transfer) return `"data": null` with `"ok": true`.

## Files

```
CCAT_MCP/
├── server.py          # FastMCP server + all four tools + OAuth token cache
├── pyproject.toml     # uv project definition + dependencies
├── uv.lock            # pinned dependency lockfile
├── requirements.txt   # pip fallback (kept in sync with pyproject)
├── .env.example
├── README.md
└── data/              # source OpenAPI spec, Postman collection, dev guide
```
