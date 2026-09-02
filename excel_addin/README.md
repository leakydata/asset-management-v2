# Cat Asset Tools — Excel add-in (.xlam)

A ribbon tab that works in every workbook. Install once, nothing to open.

Every operation is the same shape: **a sheet of rows with a Result column**. A
single asset is just a batch of one — type or paste one row and run it. There
is no separate form.

> `../excel/` is the older workbook build with the `Actions` form sheet. The two
> have diverged; this folder is the one being developed.

## The pieces

| File | Holds |
|---|---|
| `CatAssetLookup.bas` | proxy client, `=CatLookupSerial` / `=CatLookupDCN`, shared helpers, `TargetBook()` |
| `CatBatchLookup.bas` | batch lookup by serial and by DCN |
| `CatBatchOps.bas` | the three operation sheets — build, validate, run |
| `CatQuickLook.bas` | single-serial floating lookup |
| `CatRibbon.bas` | ribbon callbacks, per-user settings |
| `frmCatQuickLook_FORM_CODE.txt` | the Quick Look UserForm — **built by hand once**, see below |
| `customUI14.xml` | the ribbon tab |

`CatBatchOps` replaces both the old `CatActions` (the named-cell form) and
`CatBatchActions` (add/update only). Expire and Transfer became batch-capable
in the process — they weren't before.

## Column order is the interface

Excel pastes a copied block into **adjacent** destination columns, so the
action sheets lead with exactly the fields the lookup leads with, in the same
order. Moving data across is a drag-select and a paste — no Ctrl-clicking.

```text
batch-lookup results sheet
  A  QuerySerial / QueryDCN
  B  SerialNumber   C  MakeCode      D  DCN          E  OwnershipType
  F  Model          G  ModelYear
  H  Status         I  StatusName    J  OwnershipTypeName
  K  OwnershipRequestType             L  HasSubscription
  M  DcnName        N  CCID          O  CcidName
  P  DealerCode     Q  DealerName    R  DealerMakeCode
  S  MakeName       T  ProductFamilyCode              U  ProductFamilyName
  V  AssetName      W  BaseAssetName
```

| Copy | Paste into | Fills |
|---|---|---|
| `B:C` | `Cat Transfer` A2 | Serial, Make Code |
| `B:D` | `Cat Expire` A2 | Serial, Make Code, DCN |
| `B:G` | `Cat Add-Update` A2 | Serial, Make Code, DCN, Ownership Type, Model, Model Year |

Each action's required set is a prefix of the next, so one lookup ordering
serves all three. Columns 7+ of the lookup are reference detail, grouped:
record state, who holds it, what the asset is.

**Dealer Make Code is optional and sits last on every sheet.** The spec says
*"one of makeCode or dealerMakeCode must be provided"* — they're two separate
code systems for the same manufacturer, not two halves of one. Cat's code for
Caterpillar is `CAT`; our dealer code for it is `AA`. Use **Make Code**: it
comes back on every lookup and means the same thing to everyone. Fill Dealer
Make Code only if some source system hands you `AA` and not `CAT` — and then
clear Make Code, because supplying both is rejected.

| Sheet | Columns |
|---|---|
| `Cat Add-Update` | Serial, Make Code, DCN, Ownership Type, Model, Model Year, Product Family Code/Name, Base/Custom Asset Name, Dealer Make Code, Result |
| `Cat Expire` | Serial, Make Code, DCN, Dealer Make Code, Result |
| `Cat Transfer` | Serial, Make Code, Status, Reason, Dealer Make Code, Result |

Headers are matched by name — case, spaces and punctuation ignored — so extra
columns and reordering are harmless. `HeaderArray` in `CatAssetLookup` and
`SheetSpec` in `CatBatchOps` must stay in step; changing one alone silently
breaks the paste.

`OwnershipRequestType` is only populated on **PENDING** records: `RECEIVED`
means another dealer asked and you must approve or reject, `SENT` means you
asked and are waiting. Filter a lookup on `RECEIVED` and you have your
Transfer sheet.

`HasSubscription` is tri-state: `TRUE`, `FALSE`, or blank when the API didn't
return the field — which is not the same as false.

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

## Quick Look

One serial, a floating window, no sheet. Batch lookup already handles a single
serial but spends a whole sheet doing it — ten checks left you with ten
results sheets.

Select a cell holding a serial and click **Quick Look** and it pre-fills;
otherwise type one and press Enter. Every ownership record for that serial
appears in the top list (your DUSHORE example gives two — same CCID, two
DCNs), and picking one shows all 22 fields as field/value.

The window is **modeless** — it floats, so you can keep clicking cells while
it's open. That's what makes the two buttons work:

- **Copy Row** puts the whole record on the clipboard as two tab-separated
  lines, headers then values. Drops into Excel, an email or a ticket.
- **Paste at Selection** writes into whatever cell is selected *at the moment
  you click*, so it's never stale. It pastes **contextually** — 2 columns on
  `Cat Transfer`, 3 on `Cat Expire`, 6 on `Cat Add-Update`, all 22 anywhere
  else. Same nesting as the drag-select blocks.

So: sit on `Cat Add-Update` row 5, look a serial up, click Paste, and the six
required fields land at A5. For one-off rows that beats the drag-select.

Pasted cells are forced to text format — a serial like `00123`, or a DCN that
looks numeric, would otherwise be silently converted to a number and stop
matching CCAT.

## Validation rules

Every operation requires a Serial and a Make Code. Dealer Make Code is the
optional alternative — the API rejects both together and rejects neither.

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

1. Blank workbook → import the five `.bas` files plus `JsonConverter.bas`
   (VBA-JSON) → add the **Microsoft Scripting Runtime** reference.
   Then build the Quick Look form — `frmCatQuickLook_FORM_CODE.txt` has every
   property to type into the designer and the code to paste. **The code
   contains no sizes at all** — it never reads or writes Width, Height, Left,
   Top, Font, ColumnCount or ColumnWidths, so the form looks exactly as you
   draw it and stays that way. Skip the form entirely and everything else
   still works; only the Quick Look button will error.
2. Delete any `Config` sheet — settings live in the registry now, and you don't
   want a key inside the file.
3. **File ▸ Save As ▸ Excel Add-In (`.xlam`)**, named `CatAssetTools.xlam`.
4. Close it, open it in the
   [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor) →
   **Insert ▸ Office 2010+ Custom UI Part** → paste `customUI14.xml` → Validate → Save.
5. **File ▸ Options ▸ Add-ins ▸ Manage: Excel Add-ins ▸ Go** → tick it.
6. **CCAT ▸ Settings** — type the proxy URL and key, or point at a
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
