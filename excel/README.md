# Excel demo — serial-number lookup (via the proxy)

A worksheet function `=CatLookupSerial("<serial>")` that calls the **Asset
Management Proxy** (an Azure Function) and spills one row per matching ownership
record. The proxy holds the Cat credentials server-side, so **this workbook holds
no client secret** — only the proxy URL and a function key. Share the `.xlsm`
with anyone allowed to have the function key.

## What you'll see

Type a serial number and the function returns a spilled table:

| SerialNumber | MakeCode | MakeName | Model | ModelYear | AssetName | DealerCode | DealerName | DCN | DcnName | CCID | CcidName | OwnershipType | Status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 9303 | SB6 | SNOW WOLF | SSL | 2014 | | E480 | WHEELER | | RJT EXCAVATING INC | 2969528338 | R J T EXCAVATING INC | OWNED | ACTIVE |
| … | | | | | | | | | | | | | |

(A single serial can match several records across makes/dealers — each becomes a row.)

## Setup (one time)

1. **Open a new workbook** and save it as **Excel Macro-Enabled Workbook (`.xlsm`)**.

2. **Add a `Config` sheet** with labels in column A, your values in column B:

   | A | B |
   |---|---|
   | `ProxyUrl` | `https://<app>.azurewebsites.net/api` |
   | `FunctionKey` | _the function key from the proxy_ |
   | `PartyNumber` | _(optional — usually blank; the proxy supplies the dealer code)_ |

3. **Open the VBA editor** (`Alt`+`F11`) and import these modules
   (File ▸ Import File…):
   - `JsonConverter.bas`  (VBA-JSON — <https://github.com/VBA-tools/VBA-JSON>)
   - `CatAssetLookup.bas`  (core proxy client + the `=CatLookupSerial` function)
   - `CatBatchLookup.bas`  (the batch lookup macro)
   - `CatActions.bas`  (single add / update / expire / transfer / check macros)
   - `CatBatchActions.bas`  (the batch add / update macro)

4. **Add the Scripting Runtime reference** (Tools ▸ References…) — check
   **Microsoft Scripting Runtime**. VBA-JSON needs it for `Scripting.Dictionary`.

5. Back in Excel, in any cell:
   ```
   =CatLookupSerial("9303")
   =CatLookupDCN("12345")
   ```
   or point them at a cell: put a serial in `B1` and use `=CatLookupSerial(B1)`,
   or a DCN in `B2` and use `=CatLookupDCN(B2)`.

   `CatLookupDCN` returns the same 14-column spilled table, filtered by Dealer
   Customer Number instead of serial — one row per ownership record. A DCN
   typically has many assets, so expect a taller spill than a serial lookup;
   leave empty rows below the formula or Excel shows a `#SPILL!` error.

## Batch lookup macro

For a whole list of serials, use the macro instead of the live function:

1. Put your serial numbers in a column on any sheet (a header like "Serial" is fine).
2. Run the macro **`CatBatchLookup`** (Developer ▸ Macros, or assign it to a button/shape).
3. When prompted, **select the column of serials** and click OK.
4. It creates a new **"Asset Lookup Results"** sheet with one row per ownership
   record, plus a `QuerySerial` column (so you can tie rows back to your input)
   and a `Note` column.

Each serial gets exactly one of three outcomes, so a failed call is never
mistaken for an empty result:

| Outcome | Note value | Row shading | Meaning |
|---|---|---|---|
| Found | _(blank)_ | none | one row per ownership record |
| Not in CCAT | `NOT IN CCAT` | grey | the proxy answered successfully with **zero** records |
| Failed | `LOOKUP FAILED: …` | red | the call errored / timed out — **we don't know** if the asset exists |

`LOOKUP FAILED` rows are highlighted so you can spot and **re-run just those**.
Transient failures (timeouts, HTTP 429/5xx, connection drops) are retried
automatically up to 3 times before being marked failed; permanent errors (e.g.
400/401/403) are not retried. The macro writes values once (not a live formula),
so the results sheet won't re-call the proxy on recalculation.

> Tip: to run it from a button, insert a shape (Insert ▸ Shapes), right-click ▸
> **Assign Macro** ▸ `CatBatchLookup`.

## Actions: add / update / expire / transfer

The `CatActions` module covers the API's write operations plus a read-only
ownership check. These are **macros, not worksheet functions**, because they
change data and must never fire on recalculation.

1. Run **`CatSetupActionsSheet`** once — it builds an **"Actions"** sheet with
   labeled input cells and buttons.
2. Fill the **Asset Identifier** block (Serial, Make Code *or* Dealer Make Code,
   DCN), plus the fields for the action you want, then click its button:

   | Button | Proxy route | Notes |
   |---|---|---|
   | Add / Update | `POST /api/ownership` | New records require Ownership Type, Model, Model Year |
   | Expire | `POST /api/expire` | Removes the record (asks to confirm) |
   | Approve / Reject | `POST /api/transfer` | Status APPROVED/REJECTED; reason required to reject; no DCN |
   | Check Ownership | `GET /api/search` | Read-only; summarizes who owns the serial |

3. Each action asks for confirmation, then writes the outcome (record status, or
   the proxy's error code + message) to the **Result** cell.

> ⚠️ **Writes are gated by the proxy.** If the proxy's `CAT_ENABLE_WRITES` is
> `false` (the default), the three write buttons return **`writes_disabled`** —
> that's expected until an admin enables writes. The proxy acts on the dealer
> code it is configured with; keep that on a **test dealer code** until you
> intend to touch production. The Check button is always safe.

## Batch add / update macro

The write counterpart to `CatBatchLookup`: add or update many assets from a
sheet in one run (module `CatBatchActions`).

1. Run **`CatSetupBatchAddUpdateSheet`** once — it builds a **"Batch Add-Update"**
   sheet with the right column headers, an Ownership Type dropdown, and a
   **Run Batch Add / Update** button.
2. Fill **one row per asset**. Columns (matched by header name, so order and
   extra columns don't matter):

   | Column | Required | Notes |
   |---|---|---|
   | `Serial` | ✅ | Asset serial number (exact match) |
   | `Make Code` *or* `Dealer Make Code` | ✅ | Exactly one of the two (e.g. `CW1` / `CW`) |
   | `DCN` | ✅ | Dealer Customer Number |
   | `Ownership Type` | new records | `owned/rental/leased/sold/inventory/unknown` |
   | `Model` | new records | e.g. `980H` (max 65 chars) |
   | `Model Year` | new records | 4-digit year of manufacture, e.g. `2006` |
   | `Product Family Code` / `Product Family Name` | optional | e.g. `MDWL` / `MEDIUM WHEEL LOADER` (max 50) |
   | `Base Asset Name` | optional | Canonical name set by the dealer (max 60) |
   | `Custom Asset Name` | optional | Your own label; **shown in preference to** Base Asset Name (max 60) |
   | `Result` | — | Written by the macro (don't edit) |

   Headers are matched by name (spaces/case ignored), so the friendly labels
   above and the bare API names (`MakeCode`, `ModelYear`, …) both work. On an
   **existing** record, only the columns you fill are updated; blank optional
   cells are left untouched.

3. Click **Run** (one confirmation for the whole batch). Each row gets a
   colour-coded outcome in its `Result` cell:

   | Result | Shading | Meaning |
   |---|---|---|
   | `OK (200/201) - <status>` | green | added / updated |
   | `FAILED …` | red | the proxy/API rejected it (incl. `writes_disabled`) — safe to re-run, add/update is idempotent |
   | `SKIPPED …` | yellow | missing a required field; nothing sent |

Transient failures (timeouts, 429, the proxy's 5xx) are retried up to 3 times.

> ⚠️ **Same write gate as the buttons.** With the proxy's `CAT_ENABLE_WRITES`
> `false` (default), every row returns `403 writes_disabled` and nothing changes.
> Keep the proxy on a **test dealer code** until you intend to touch production.

## Notes

- **No credentials in the workbook.** Only `ProxyUrl` + `FunctionKey` live here.
  Treat the function key like a password — but it is revocable on the proxy
  without touching the Cat credentials.
- **Which dealer code?** The proxy decides the `partyNumber` server-side. Leave
  `Config!PartyNumber` blank unless an admin tells you otherwise.
- **Errors are explicit.** Failed calls surface the proxy envelope's
  `error.code` + `error.description`, so a failure is never mistaken for an empty
  result.
- **Recalculation:** as a live function, Excel re-runs `CatLookupSerial` when its
  input changes (and on full recalc). Each run is one proxy call. For many
  serials, prefer the `CatBatchLookup` macro, which writes values once.
- `CatLookupSerial` filters by `serialNumber`; `CatLookupDCN` filters by `dcn`.
  Both are exact-match (the API has no partial/fuzzy search) and both run
  through the same `CatSearch` building block.
