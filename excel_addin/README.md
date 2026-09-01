# Cat Asset Tools — Excel add-in (.xlam) build

The same five modules as [`../excel/`](../excel/), converted to run as a loaded
**add-in** rather than as code inside one workbook. Install it once and the
buttons work in every workbook, with nothing to open.

Use this folder if you want a ribbon tab. Use `../excel/` if you want a
self-contained `.xlsm` someone can email.

## What changed, and why

**`ThisWorkbook` is the add-in.** That's the whole conversion. In an `.xlam`,
`ThisWorkbook` is the hidden add-in file itself — so the original code would
have built the Actions sheet, the Batch Add-Update sheet and every results
sheet *inside a workbook nobody can see*. Every button would have appeared to
run and produced nothing.

The rule applied throughout: anything the user is meant to see goes to the
workbook they're working in; only genuine add-in resources stay `ThisWorkbook`.

| Where | Before | After |
|---|---|---|
| `CatActions.NV` | `ThisWorkbook.Names` | `TargetBook().Names` |
| `CatActions.SetResult` | `ThisWorkbook.Names` | `ActiveWorkbook.Names` (never raises — it runs inside error handlers) |
| `CatActions.CatSetupActionsSheet` | `ThisWorkbook.Worksheets` | `TargetBook().Worksheets` |
| `CatActions.AddList` | `ThisWorkbook.Names` | `ws.Parent.Names` |
| `CatBatchActions.TargetSheet` | `ThisWorkbook.Worksheets` | `TargetBook().Worksheets` |
| `CatBatchActions.CatSetupBatchAddUpdateSheet` | `ThisWorkbook.Worksheets` | `TargetBook().Worksheets` |
| `CatBatchLookup.MakeOutputSheet` | `ThisWorkbook.Worksheets` | takes a `Workbook` argument |
| Shape `OnAction` | bare macro name | `'CatAssetTools.xlam'!Macro` |

`TargetBook()` is new, in `CatAssetLookup`. It returns `ActiveWorkbook` and
raises a readable error if there's no workbook open, or if the add-in somehow
targets itself — otherwise those cases fail as a confusing "required field"
message much later.

Results from a batch lookup go to `rng.Worksheet.Parent` — the workbook the
selection came from — rather than `ActiveWorkbook`, so they land next to the
input even if the user picked a range in another workbook.

**Config moved out of the worksheet.** The old `Config` sheet would ship
*inside* the distributed `.xlam`, function key and all. `Cfg()` now reads
per-user registry settings:

```vb
Private Function Cfg(ByVal label As String) As String
    Cfg = GetSetting(REG_APP, REG_SECTION, label, "")
End Function
```

The labels are unchanged — `ProxyUrl`, `FunctionKey`, `PartyNumber` — so no
call site needed editing. `CatRibbon.CatSettings` prompts for all three and is
wired to the ribbon's Settings button. `CatClearSettings` wipes them, which
matters on a shared machine.

This makes `CatAssetLookup` depend on `CatRibbon` for the `REG_APP` /
`REG_SECTION` constants. Import both.

**Ribbon callbacks.** A ribbon `onAction` must be
`Sub Name(control As IRibbonControl)`, and the existing macros take no
arguments. `CatRibbon` holds ten one-line wrappers, so `CatAddOwnership` and
friends stay callable from Alt+F8 and Assign Macro exactly as before.

## Build it

1. Open a blank workbook, import all five `.bas` files plus `JsonConverter.bas`
   (VBA-JSON), and add the **Microsoft Scripting Runtime** reference
   (Tools ▸ References).
2. Delete any `Config` sheet — it's not used any more and you don't want a key
   in the file.
3. **File ▸ Save As ▸ Excel Add-In (`.xlam`)**. Name it `CatAssetTools.xlam`;
   it defaults to `%APPDATA%\Microsoft\AddIns\`.
4. Close it in Excel, then open the `.xlam` in the
   [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor) →
   **Insert ▸ Office 2010+ Custom UI Part** → paste `customUI14.xml` → Validate
   → Save.
5. Back in Excel: **File ▸ Options ▸ Add-ins ▸ Manage: Excel Add-ins ▸ Go** and
   tick it.
6. Click **Cat Assets ▸ Settings** and enter the proxy URL and function key.

## Deploying to other people

Put the `.xlam` on a share and add that share as a **Trusted Location**
(File ▸ Options ▸ Trust Center ▸ Trust Center Settings ▸ Trusted Locations,
with *Allow Trusted Locations on my network* ticked). Without it Excel blocks
the macros and the buttons silently do nothing.

Each person installs it once and enters their own key via Settings. Nothing
sensitive travels with the file.

## Known gotchas

- **Macros don't appear in Alt+F8** for a loaded add-in. That's normal — the
  ribbon is how you run them; type the name if you need the dialog.
- **`=CatLookupSerial(...)` needs the add-in properly installed** via the
  Add-ins dialog. If you merely *open* the `.xlam`, Excel writes the full path
  into the formula.
- **Blank button icon** means an `imageMso` name is wrong. Cosmetic; swap it.
- **"Expected variable or procedure, not module"** — the module
  `CatBatchLookup` also contains a `Sub CatBatchLookup`, and in a compile-time
  reference the module name wins. `CatRibbon` calls it module-qualified as
  `CatBatchLookup.CatBatchLookup`. The shape `OnAction` path never hit this
  because a macro-name string resolves at run time, not compile time.
- **"User-defined type not defined"** on compile means the Office object
  library reference is missing — change `IRibbonControl` / `IRibbonUI` to
  `Object` in `CatRibbon`, or add the reference.

## What this still doesn't fix

The add-in talks to the Azure proxy directly, so none of it writes a
`CCAT_AUDIT` row and none of the Snowflake business rules apply. Manual edits
are the ones most worth having a record of. See the Streamlit-in-Snowflake
option before this becomes the permanent home for manual updates.
