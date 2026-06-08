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

3. **Open the VBA editor** (`Alt`+`F11`) and import both modules
   (File ▸ Import File…):
   - `JsonConverter.bas`  (VBA-JSON — you already have it)
   - `CatAssetLookup.bas`  (this folder)

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
   and a `Note` column (`no records found` / `ERROR: …` where applicable).

The macro reuses one cached token for the entire batch and shows progress in the
status bar. Because it writes values once (not a live formula), the results sheet
won't re-call the API on recalculation.

> Tip: to run it from a button, insert a shape (Insert ▸ Shapes), right-click ▸
> **Assign Macro** ▸ `CatBatchLookup`.

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
