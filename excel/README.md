# Excel demo — serial-number lookup

A worksheet function `=CatLookupSerial("<serial>")` that calls the Cat Asset
Management V2 API and spills one row per matching ownership record.

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
   | `ClientId` | _your client id_ |
   | `ClientSecret` | _your client secret_ |
   | `Scope` | _your client id_`/.default` |
   | `TenantId` | `ceb177bf-013b-49ab-8a9c-4abce32afc1e` |
   | `PartyNumber` | `ZZIO` |
   | `BaseUrl` | `https://services.cat.com/catDigital/assetManagement/v2` |
   | `TokenUrl` | _(leave blank)_ |

3. **Open the VBA editor** (`Alt`+`F11`) and import these modules
   (File ▸ Import File…):
   - `JsonConverter.bas`  (VBA-JSON — you already have it)
   - `CatAssetLookup.bas`  (core API + the `=CatLookupSerial` function)
   - `CatBatchLookup.bas`  (the batch lookup macro)
   - `CatActions.bas`  (add / update / expire / transfer / check macros)

4. **Add the Scripting Runtime reference** (Tools ▸ References…) — check
   **Microsoft Scripting Runtime**. VBA-JSON needs it for `Scripting.Dictionary`.

5. Back in Excel, in any cell:
   ```
   =CatLookupSerial("9303")
   ```
   or point it at a cell: put a serial in `B1` and use `=CatLookupSerial(B1)`.

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
| Not in CCAT | `NOT IN CCAT` | grey | the API answered successfully with **zero** records |
| Failed | `LOOKUP FAILED: …` | red | the call errored / timed out — **we don't know** if the asset exists |

`LOOKUP FAILED` rows are highlighted so you can spot and **re-run just those**.
Transient failures (timeouts, HTTP 429/5xx, connection drops) are retried
automatically up to 3 times before being marked failed; permanent errors (e.g.
403) are not retried. The macro reuses one cached token for the whole batch,
shows progress in the status bar, and writes values once (not a live formula),
so the results sheet won't re-call the API on recalculation.

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

   | Button | Endpoint | Notes |
   |---|---|---|
   | Add / Update | `POST /ownershipRecords` | New records require Ownership Type, Model, Model Year |
   | Expire | `POST /ownershipRecords/expire` | Removes the record (asks to confirm) |
   | Approve / Reject | `POST /ownershipRequests/transfer` | Status APPROVED/REJECTED; reason required to reject; no DCN |
   | Check Ownership | `POST /ownershipRecords/search` | Read-only; summarizes who owns the serial |

3. Each action asks for confirmation, then writes the outcome (record status, or
   the API's error code + message) to the **Result** cell.

> ⚠️ **These change real data.** Add/expire/transfer act on the dealer code in
> `Config!PartyNumber`. Keep that on your **test dealer code** and use test
> assets until you intend to touch production. The Check button is always safe.

## Notes

- **Credentials live in the `Config` sheet, not in code** — so don't share the
  `.xlsm` with anyone who shouldn't have the client secret. Treat the file like
  a password.
- **The token is cached** in memory and auto-refreshes ~1 minute before expiry,
  so repeated lookups don't re-authenticate every time. Use the
  `CatClearTokenCache` macro to force a fresh token (e.g. after changing creds).
- **Recalculation:** as a live function, Excel re-runs `CatLookupSerial` when its
  input changes (and on full recalc). Each run is one API call. For a heavy demo
  with many serials, consider copying results as values, or ask and I can supply a
  button-driven macro version that writes results once instead of recalculating.
- **Dealer code:** this demo searches as `PartyNumber = ZZIO` (your entitled test
  dealer code). Change the `PartyNumber` cell to your production code once
  Caterpillar upgrades your credentials.
- The function uses `serialNumber` as the filter. To also support DCN lookups,
  the underlying `CatSearch` already accepts a `dcn` argument — say the word and
  I'll expose a `=CatLookupDcn(...)` too.
```
