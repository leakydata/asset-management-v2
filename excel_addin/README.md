# Cat Asset Tools — Excel add-in (.xlam)

A ribbon tab that works in every workbook. Install once, nothing to open.

Every operation is the same shape: **a sheet of rows with a Result column**. A
single asset is just a batch of one — type or paste one row and run it. There
is no separate form.

> `../excel/` is the older workbook build with the `Actions` form sheet. The two
> have diverged; this folder is the one being developed.

## The four modules

| Module | Holds |
|---|---|
| `CatAssetLookup.bas` | proxy client, `=CatLookupSerial` / `=CatLookupDCN`, shared helpers, `TargetBook()` |
| `CatBatchLookup.bas` | batch lookup by serial and by DCN |
| `CatBatchOps.bas` | the three operation sheets — build, validate, run |
| `CatRibbon.bas` | ribbon callbacks, per-user settings |

`CatBatchOps` replaces both the old `CatActions` (the named-cell form) and
`CatBatchActions` (add/update only). Expire and Transfer became batch-capable
in the process — they weren't before.

## Column order is deliberate

The leading columns of every operation sheet appear in the **same relative
order as the batch-lookup output**, so you can Ctrl-select the matching columns
on a results sheet, copy, and paste them straight in. Excel pastes a
multi-area column selection as one contiguous block in left-to-right order.

```text
lookup output (22 columns)
  1 SerialNumber     2 MakeCode        3 MakeName         4 Model
  5 ModelYear        6 AssetName       7 DealerCode       8 DealerName
  9 DCN             10 DcnName        11 CCID            12 CcidName
 13 OwnershipType   14 Status         15 OwnershipTypeName
 16 StatusName      17 HasSubscription
 18 OwnershipRequestType               19 DealerMakeCode
 20 ProductFamilyCode                  21 ProductFamilyName
 22 BaseAssetName

Ctrl-select these ten, in this order, and paste into Add-Update column 1:
  1 Serial   2 MakeCode   4 Model   5 ModelYear   9 DCN   13 OwnershipType
 19 DealerMakeCode  20 ProductFamilyCode  21 ProductFamilyName  22 BaseAssetName
       ->  Add-Update columns 1..10
```

The first 14 positions are frozen — new fields are appended, never inserted —
so anything already built against the old layout keeps working.

**Take only one make column.** `MakeCode` and `DealerMakeCode` both come back
populated, but supplying both is rejected. Paste all ten and then clear
whichever you don't want, or Ctrl-select just the first six.

`OwnershipRequestType` is only populated on **PENDING** records: `RECEIVED`
means another dealer asked and you must approve or reject, `SENT` means you
asked and are waiting. Blank means no pending request, or one you can't act
on. Filter on it to build the Transfer sheet.

`HasSubscription` is tri-state: `TRUE`, `FALSE`, or blank when the API didn't
return the field at all — which is not the same as false.

| Sheet | Columns |
|---|---|
| `Cat Add-Update` | Serial, Make Code, Model, Model Year, DCN, Ownership Type, Dealer Make Code, Product Family Code/Name, Base/Custom Asset Name, Result |
| `Cat Expire` | Serial, Make Code, DCN, Dealer Make Code, Result |
| `Cat Transfer` | Serial, Make Code, Dealer Make Code, Status, Reason, Result |

Anything with no lookup counterpart sits after the aligned block. Headers are
matched by name — case, spaces and punctuation ignored — so extra columns and
reordering are both harmless.

## Two buttons drive all three

**Validate** and **Run** act on the **active sheet** and work out the operation
from its name, falling back to its headers if you've renamed it (it says which
operation it inferred, and waits for you to confirm).

- **Validate** sends nothing. It checks every row and writes the outcome to the
  Result column — green ready, yellow skipped with the reason. On Add/Update it
  also lists which optional fields are being left out, so the omit-blanks rule
  is visible before anything goes.
- **Run** confirms once with the row count, then sends. Green ok, red failed,
  yellow skipped.

A batch-lookup **results** sheet also has an `OwnershipType` column, so it
would otherwise look like an Add-Update sheet and Run would fire writes for
every row. Its `QuerySerial` / `QueryDCN` column is the tell, and those sheets
are refused outright.

## Validation rules

Every operation requires a Serial and exactly one of Make Code / Dealer Make
Code — the API rejects both together and rejects neither.

| Operation | Also required |
|---|---|
| Add / Update | DCN, **Ownership Type** (one of owned/rental/leased/sold/inventory/unknown). Model Year must be 4 digits if given. |
| Expire | DCN |
| Transfer | Status = APPROVED or REJECTED; a Reason when rejecting |

**Ownership Type is required on every row by choice, not by the API.** Cat only
demands it for a *new* record. Enforcing it turns the most common rejection
into a skipped row you can see and fix. The trade: every update then rewrites
`ownershipTypeCode`, so pasting a stale value overwrites a correct one. Pasted
from a fresh lookup it's the current value, which is harmless.

Blank optional cells are still omitted from the request entirely — they never
overwrite what CCAT already holds.

## Build it

1. Blank workbook → import the four `.bas` files plus `JsonConverter.bas`
   (VBA-JSON) → add the **Microsoft Scripting Runtime** reference.
2. Delete any `Config` sheet — settings live in the registry now, and you don't
   want a key inside the file.
3. **File ▸ Save As ▸ Excel Add-In (`.xlam`)**, named `CatAssetTools.xlam`.
4. Close it, open it in the
   [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor) →
   **Insert ▸ Office 2010+ Custom UI Part** → paste `customUI14.xml` → Validate → Save.
5. **File ▸ Options ▸ Add-ins ▸ Manage: Excel Add-ins ▸ Go** → tick it.
6. **Cat Assets ▸ Settings** — type the proxy URL and key, or point at a
   two-column range (an old Config sheet works as-is).

Ribbon XML is read only when the add-in loads, so **restart Excel** after any
edit to it.

## Deploying to other people

Put the `.xlam` on a share and add that share as a **Trusted Location**
(with *Allow Trusted Locations on my network* ticked). Without it Excel blocks
the macros and the buttons silently do nothing. Each person installs once and
enters their own key.

## Known gotchas

- **Macros don't appear in Alt+F8** for a loaded add-in. Normal — use the ribbon.
- **`=CatLookupSerial(...)` needs the add-in properly installed** via the Add-ins
  dialog. If you merely *open* the `.xlam`, Excel writes the full path into the formula.
- **"Expected variable or procedure, not module"** — the module `CatBatchLookup`
  also contains a `Sub CatBatchLookup`, and in a compile-time reference the module
  name wins. `CatRibbon` calls it module-qualified as `CatBatchLookup.CatBatchLookup`.
- **"User-defined type not defined"** — the Office object library reference is
  missing; change `IRibbonControl` / `IRibbonUI` to `Object`, or add the reference.
- **Blank button icon** — an `imageMso` name is wrong. Cosmetic.

## What this still doesn't fix

The add-in talks to the Azure proxy directly, so nothing here writes a
`CCAT_AUDIT` row and none of the Snowflake business rules apply. Manual edits
are the ones most worth having a record of.
