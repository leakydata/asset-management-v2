# Cat Asset Console — single-file web client

`cat-asset-console.html` is one self-contained HTML file (no build, no
dependencies, no CDN) that exposes the same proxy routes the Excel workbook
uses: lookup, batch lookup, add/update, expire, transfer, and batch add/update.

It is the browser equivalent of `excel/` — same request bodies, same retry
rules, same three-outcome batch reporting. Anything the workbook can do against
the proxy, this can do.

## Why it can't just be double-clicked

**A browser is not WinHTTP.** The VBA modules call the proxy through
`WinHttp.WinHttpRequest.5.1`, which ignores CORS entirely. A web page cannot —
the browser refuses to hand your JavaScript a cross-origin response unless the
server says the origin is allowed. Two consequences:

1. **Opening the file directly (`file://…`) will not work.** A `file://` page
   sends `Origin: null`, which Azure's CORS list cannot allow. Serve it over
   `http://localhost` instead — from this folder:

   ```
   python -m http.server 8080
   ```
   then open <http://localhost:8080/cat-asset-console.html>.

2. **The Function App must allow that origin.** In the portal:
   *Function App → API → CORS →* add `http://localhost:8080`. Azure answers the
   preflight itself; no code change is needed.

Until that entry exists, every call fails with a network error. The
**Test connection** button detects this case specifically and tells you which
origin to add.

## Setup

1. Serve the folder and open the page (above).
2. Paste the **Proxy URL** (`https://<app>.azurewebsites.net/api`) and the
   **function key** — the same two values as the workbook's `Config` sheet.
3. Click **Test connection**. It calls `/search` with no parameters, which the
   proxy rejects with its own error envelope — enough to prove reachability,
   CORS and the key without touching an asset.

## The function key

The key is **never written into the file**. You type it at runtime, and with
*Remember on this machine* ticked it is kept in that browser's `localStorage` —
local to the browser profile, never transmitted anywhere but the proxy.

Treat the file as shareable and the key as a password. This is the same trust
model as the `.xlsm`, with one improvement: the key is not embedded in the
artifact you hand to someone else.

Because a browser cannot send a custom `x-functions-key` header without
triggering a CORS preflight on every call, the page passes the key as the
`?code=` query parameter instead. Same authentication, one fewer round trip,
and `GET /search` stays a "simple" request. Note that this puts the key in the
proxy's request logs — if that matters, switch to the header and allow the
`x-functions-key` header in the Function App's CORS settings.

## Write routes

The four write buttons post to `ownership`, `expire` and `transfer` — the routes
the Excel README documents. **The proxy source committed in `azure-function/`
does not implement them.** It has exactly one route (`search`) and no
`CAT_ENABLE_WRITES` setting, so either the deployed app is ahead of this repo or
the write half was never built.

Check before trusting the write tabs:

| Proxy answers | Meaning |
|---|---|
| `404` | The route does not exist — the deployed proxy is the read-only build |
| `403 writes_disabled` | Routes exist, writes are switched off (expected default) |
| `200` / `201` / `204` | Writes are live and you just changed production data |

Reads work either way. The page tolerates both response shapes — the
`{ ok, data, error }` envelope the VBA parses, and a bare Cat passthrough.

## Behaviour worth knowing

- **Blank fields are omitted, never sent as `""`.** Cat updates "only the fields
  provided in the request body", so an empty string would overwrite a stored
  value with a blank. The **Preview body only** button shows exactly what will be
  sent and lists what was left out.
- **Retries** cover timeouts, `429` and `5xx` only — up to 3 attempts. A `400`
  or `403` fails once, because it would fail identically again.
- **Batch lookup outcomes** are three-way, so a failure is never mistaken for an
  empty result: found (plain), `NOT IN CCAT` / `NO ACTIVE RECORDS` (grey), and
  `LOOKUP FAILED` (red — unknown, re-run just those).
- **Batch writes run at 2 in flight** regardless of the lookup concurrency
  setting, and **Validate only** shows every row's body without sending
  anything.
- Results copy as TSV straight into Excel, or download as CSV.

## Not published as a hosted page

This has to stay a local file. A page hosted on a third-party origin would
require allow-listing that origin on the Function App — letting anything served
from it call your proxy — and pasting a company function key into it.
