# Cat Asset Tools — Excel add-in (.xlam)

A ribbon tab that works in every workbook. Install once, nothing to open.

**Reading** is a floating window or a results sheet. **Writing** is always the
same shape: a sheet of rows with a Result column, checked by Validate before
Run sends anything. A single asset is just a batch of one.

> `../excel/` is the older workbook build with the `Actions` form sheet. The two
> have diverged; this folder is the one being developed.

## What's on the tab

| Group | Button | Does |
|---|---|---|
| Lookup | **Serial Lookup** | one serial in a floating window |
| | **DCN Lookup** | everything one customer owns |
| | **Batch Serials** | a column of serials → results sheet |
| | **Batch DCNs** | a column of DCNs → results sheet |
| Reconcile | **Compare to CCAT** | a NAXT sheet → only the rows that disagree |
| Build Sheet | **Add / Update**, **Expire**, **Transfer** | create an operation sheet |
| | **Undo a Run...** | rebuild what a past Run overwrote |
| Active Sheet | **Validate** | check every row, say what would change. Sends nothing |
| | **Run** | send it. Esc stops, and it is all logged |
| Setup | **Settings** | proxy URL and key, per user |
| | **Test Connection** | which part of the setup is wrong, in plain English |
| | **Logs** | what was searched, what was written, what it was before |

Right-clicking a cell also offers *Cat: Look up this serial* and *Cat: Look up
this DCN*.

Everything that reads is safe to press. **Run is the only button that writes to
CCAT** — Validate, Compare to CCAT and Undo a Run all stop at producing a sheet
for you to check first.

## The pieces

| File | Holds |
|---|---|
| `CatAssetLookup.bas` | proxy client, `=CatLookupSerial` / `=CatLookupDCN`, shared helpers, `TargetBook()` |
| `CatBatchLookup.bas` | batch lookup by serial and by DCN |
| `CatBatchOps.bas` | the three operation sheets — build, validate, run |
| `CatSerialLook.bas` | single-serial floating lookup |
| `CatDcnLook.bas` | single-DCN floating lookup — everything one customer owns |
| `CatLookShared.bas` | the paste rules both lookup windows obey |
| `CatAudit.bas` | the lookup log, the write log, and the before-images Undo reads |
| `CatRibbon.bas` | ribbon callbacks, per-user settings, right-click menu |
| `frmCatQuickLook_FORM_CODE.txt` | the Serial Lookup UserForm — **built by hand once**, see below |
| `frmCatDcnLook_FORM_CODE.txt` | the DCN Lookup UserForm — **cloned from the one above** |
| `customUI14.xml` | the ribbon tab |
| `Install.bat` / `Uninstall.bat` | what you ship in the zip - see *Deploying* |
| `BACKLOG.md` | what was built and why, what was dropped and why |

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

  It also **fetches the current record and says what would change**:

  ```text
  OK to send - CHANGES: OwnershipType RENTAL -> OWNED
  OK to send - no change
  OK to send - NEW record for this DCN
  ```

  That closes the footgun below: every update rewrites `ownershipTypeCode`, so
  a sheet pasted from last week silently reverts a record someone has fixed
  since, and nothing used to warn you.

  The diff is **advisory only** — it never turns an OK row into SKIPPED, because
  a row's fate must not depend on a live call, and Validate has to keep working
  when the proxy is down. It costs one lookup per distinct serial; sheets over
  50 rows ask first.

- **Run** confirms once with the row count, then sends. Green ok, red failed,
  yellow skipped.

  **Esc stops it** between rows, never mid-request. A stopped run says
  `STOPPED at row 312`, and re-running offers to skip the rows that already
  came back OK — so an interrupted 500-row batch finishes without re-sending
  the first 299.

  **Every row it sends is logged**, with the state of the record *before* it was
  touched. See *Logs* and *Undo a Run*.

A batch-lookup **results** sheet also has an `OwnershipType` column, so it
would otherwise look like an Add-Update sheet and Run would fire writes for
every row. Its `QuerySerial` / `QueryDCN` column is the tell, and those sheets
are refused outright.

## Serial Lookup

One serial, a floating window, no sheet. Batch lookup already handles a single
serial but spends a whole sheet doing it — ten checks left you with ten
results sheets.

Select a cell holding a serial and click **Serial Lookup** and it pre-fills;
otherwise type one and press Enter. Every ownership record for that serial
appears in the top list (your DUSHORE example gives two — same CCID, two
DCNs), and picking one shows all 22 fields as field/value.

The box remembers the **last 25 things you searched for** — type, or pick from
the dropdown. Serials and DCNs keep separate lists, and only searches that
found something are remembered, so typos never silt it up.

The window is **modeless** — it floats, so you can keep clicking cells while
it's open. That's what makes the four buttons work:

- **Copy Row** puts the whole record on the clipboard as two tab-separated
  lines, headers then values. Drops into Excel, an email or a ticket.
- **Paste at Selection** writes into whatever cell is selected *at the moment
  you click*, so it's never stale. It pastes **contextually** — 2 columns on
  `Cat Transfer`, 3 on `Cat Expire`, 6 on `Cat Add-Update`, all 22 anywhere
  else. Same nesting as the drag-select blocks.
- **Paste All** does the same for *every* record at once. A DCN returning 55
  assets is 55 clicks otherwise. It confirms first and names the exact range,
  including how many non-empty cells it would overwrite — it is the only thing
  either window does that writes over a block rather than a row.
- **Send List to Sheet** writes the whole result list to a new sheet, shaped
  like a batch results sheet. That shape is deliberate: it carries a `Query*`
  column, so Validate and Run refuse it and a reference sheet can never be
  mistaken for an operation sheet.

You can also **right-click any cell** and choose *Cat: Look up this serial* or
*Cat: Look up this DCN*, which seeds the window from that cell.

So: sit on `Cat Add-Update` row 5, look a serial up, click Paste, and the six
required fields land at A5. For one-off rows that beats the drag-select.

Pasted cells are forced to text format — a serial like `00123`, or a DCN that
looks numeric, would otherwise be silently converted to a number and stop
matching CCAT.

## DCN Lookup

The same window, asking the opposite question. Serial Lookup asks *who owns
this machine*; DCN Lookup asks *what does this customer own*. `CatSearch`
already took both arguments, so it's the same call with the other one filled.

Two things differ, both because the question inverts. A serial returns two or
three ownership records; DCN `C00103757` returns **55 assets**.

- **The list shows the asset** — serial, make, model, year, type, status —
  because DCN and DCN Name would repeat identically down all 55 rows and spend
  the two widest columns saying nothing. The customer name goes on the title
  bar, where it's said once.
- **The record list is taller** — fourteen rows against six. It takes the slack
  the detail list wasn't using; that one still shows all 22 fields without
  scrolling.

All four buttons behave identically to the Serial Lookup window, and
deliberately share their column rules via `CatLookShared` — pasting onto
`Cat Expire` fills the same three columns from either window. **Paste All** is
the one that earns its place here: 55 assets is 55 clicks otherwise.

## Compare to CCAT

The reason most people open this add-in: check serials, add the missing ones,
correct the ones CCAT has wrong.

Put a sheet of NAXT values in front of it — headers in row 1 — and press
**Compare to CCAT**. It reads whatever columns it recognises, looks each serial
up, and writes **only the rows CCAT disagrees with**, already in Add/Update
shape and ready to Validate and Run. Rows that already match never appear;
reviewing those is the work being removed.

It sends nothing, and Esc stops it.

### The columns it recognises

Header matching ignores case, spaces, underscores and punctuation, and knows
the names real exports use:

| Field | Also accepted as | Role |
|---|---|---|
| Serial | `SERIALNUMBER`, `SERIALNO`, `ASSETSERIAL` | match key |
| DCN | `CUSTOMERNUMBER`, `CUSTNUMBER`, `DEALERCUSTOMERNUMBER` | match key |
| Make Code | `THREE_DIGIT_CAT_MAKE_CODE`, `MAKE` | compared |
| Model Year | `MANUFACTURERYEAR` | compared |
| Ownership Type | `OWNERSHIPTYPECODE` | compared |
| Model | — | written on NEW records only |

**`ENGINESERIALNUMBER`, `TRANSMISSIONSERIALNUMBER`, `GENERATORSERIALNUMBER` and
`DEALEREQUIPMENTSERIALNUMBER` are deliberately ignored.** A real export carries
several of those beside the real one, and matching the wrong column would look
up serials CCAT has never heard of and report the whole sheet as missing.

Only cells that are **filled in** get compared. A blank is omitted from an
Add/Update request entirely and cannot change anything, so calling it a
difference would be a lie.

### Model is not compared, on purpose

D365 and Cat spell the same machine differently — `303.5ECR` against
`303.5E2CR`. On a ten-row sample, five rows matched their CCAT record on DCN
and differed on **nothing but** that. Cat's spelling is the authoritative one,
so comparing Model would mean proposing to overwrite Cat's own model name on
every machine.

Those rows are not thrown away: they go to a **`Cat Model Mismatches`** sheet
listing the serial, Cat's model and yours. That is a list to correct **D365**
from, pointed the right way. It carries a `QuerySerial` column and no Result
column, so it cannot be Validated or Run.

### Picking the right record

A serial can sit on several ownership records under different DCNs, so the
right one is chosen and never guessed:

| Situation | What happens |
|---|---|
| sheet has a DCN column | match on it — definitive |
| exactly one record | that one |
| several records, no DCN | **AMBIGUOUS** |

Ambiguous rows land on the sheet with the DCN blank, so Validate marks them
yellow with a reason — a worklist of rows needing a DCN, rather than a number
in a dialog. **A NAXT export really wants a DCN column**; without one, only
single-record serials and brand-new ones can be resolved.

A blank DCN coming *back* from CCAT means something else entirely: Cat returns
the DCN only on our own dealer's records, so blank means **another dealer holds
that machine**.

## Logs

Two CSVs a month under `%APPDATA%\CatAssetTools`, opened by **CCAT > Logs**.

| File | Holds |
|---|---|
| `cat-lookup-log-YYYY-MM.csv` | every serial and DCN searched, from either window or either batch |
| `cat-write-log-YYYY-MM.csv` | every row Run sent, **plus the full 22-field state of the record before it was touched** |

Kept apart on purpose: a few hundred searches would bury the handful of writes
that matter, and the write log has to stay usable as an audit trail.

Worksheet-function calls are not logged — `=CatLookupSerial()` re-evaluates on
any recalc, and that is Excel, not a person.

Every field is quoted. A DCN name like `SMS RENTAL (WA) PTY LTD, INC` would
otherwise shift every later column, and a log that misreports which record was
touched is worse than no log.

## Undo a Run

**Build Sheet > Undo a Run...** lists recent runs and rebuilds what one
overwrote, from the before-images in the write log.

**It sends nothing.** You get a `Cat Undo <runid>` sheet, and you Validate and
Run it like any other work — an undo that sends is a second unreviewed write on
top of the first, at the moment the person is least likely to be checking.

Three cases, kept apart because they are genuinely different:

| Before-state | What happened | What Undo does |
|---|---|---|
| `FOUND` | record changed or expired | restores the old values |
| `NONE` | the run **created** it | offers a separate Expire sheet, behind its own confirmation |
| `UNAVAILABLE` | the lookup failed first | **reported, never restored** |

That last row matters most: restoring all-blanks because a lookup timed out
would write emptiness over a live record. Rows that **failed** at the time are
skipped too — they changed nothing.

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

1. Blank workbook → import the eight `.bas` files plus `JsonConverter.bas`
   (VBA-JSON) → add the **Microsoft Scripting Runtime** reference.
   Then build the Serial Lookup form — `frmCatQuickLook_FORM_CODE.txt` lists the
   controls to drop on and the code to paste. **Drop them anywhere and name
   them**; the code sets the form's size, stretches the full-width controls to
   the edges, and sets both lists' `ColumnCount`/`ColumnWidths` and the column
   heading. Each `Top` and each `Font` is still yours in the designer.
   Then clone it for DCN Lookup — `frmCatDcnLook_FORM_CODE.txt` has the
   export/rename/import steps. Skip either form and everything else still
   works; only that form's button will error.
2. Delete any `Config` sheet — settings live in the registry now, and you don't
   want a key inside the file.
3. **File ▸ Save As ▸ Excel Add-In (`.xlam`)**, named `CatAssetTools.xlam`.
4. Close it, open it in the
   [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor) →
   **Insert ▸ Office 2010+ Custom UI Part** → paste `customUI14.xml` → Validate → Save.
5. **File ▸ Options ▸ Add-ins ▸ Manage: Excel Add-ins ▸ Go** → tick it.
6. **CCAT ▸ Settings** — type the proxy URL and key, or point at a
   two-column range (an old Config sheet works as-is). It offers to test the
   connection when you save; say yes.
7. **CCAT ▸ Test Connection** any time after that. It sends one harmless read
   for a serial that deliberately does not exist, so it proves the URL, the key
   and the network without depending on any asset still being in CCAT, and it
   names which part is wrong in plain English rather than a status code.

Ribbon XML is read only when the add-in loads, so **restart Excel** after any
edit to it.

## Deploying to other people

### A zip with an installer (preferred)

Ship a folder containing exactly three files and zip it:

```text
CatAssetTools/
├── ASSET_MANAGEMENT_ADDIN.xlam    ← the built add-in
├── Install.bat
└── Uninstall.bat
```

They unzip it and double-click `Install.bat`. No admin rights — everything is
under their own profile.

**Never put the function key in the zip.** It lives per-user in `HKCU` and is
entered once via **CCAT ▸ Settings**. A key inside a file that gets forwarded
around is a key you have to rotate.

`Install.bat` does three things, and the third is the one people miss:

1. **Copies to `%APPDATA%\Microsoft\AddIns`.** That's Excel's per-user add-in
   folder *and* one of its default Trusted Locations, so the macros run without
   a prompt. Anywhere else and they get "macros have been disabled" with no
   ribbon.
2. **Drops the Mark-of-the-Web.** A file that arrived in a zip by email or
   download carries a `Zone.Identifier` stream that makes Excel refuse to load
   it. `COPY` writes only the primary stream, so the block goes with it.
3. **Registers it.** Copying alone only makes it *appear*, unticked, in the
   Add-ins dialog. Excel loads what's listed under
   `HKCU\Software\Microsoft\Office\<ver>\Excel\Options` as `OPEN`, `OPEN1`,
   `OPEN2`… — the installer finds the first free slot and writes
   `"ASSET_MANAGEMENT_ADDIN.xlam"`, quotes included, exactly as Excel does.

It refuses to run while Excel is open, for two reasons that both produce a
confusing half-install: a loaded add-in is locked so the copy fails, and Excel
rewrites its `Options` key on exit, so a registration made while it's running
is overwritten on close.

Re-running is safe — it overwrites the file and won't add a second registration.
`Uninstall.bat` removes the file and only its own `OPEN` slot, leaving other
people's add-ins and the saved settings alone.

> **Tell them to unzip first.** Double-clicking `Install.bat` from inside the
> zip viewer fails: Windows extracts just the `.bat` to a temp folder and the
> `.xlam` isn't next to it. The installer detects this and says so.

### From a network share

Alternative if you'd rather not hand out files: put the `.xlam` on a share and
add that share as a **Trusted Location** (with *Allow Trusted Locations on my
network* ticked). Without it Excel blocks the macros and the buttons silently
do nothing. Each person still installs once and enters their own key.

## Known gotchas

- **Macros don't appear in Alt+F8** for a loaded add-in. Normal — use the ribbon.
- **`=CatLookupSerial(...)` needs the add-in properly installed** via the Add-ins
  dialog. If you merely *open* the `.xlam`, Excel writes the full path into the formula.
- **"Expected variable or procedure, not module"** — the module `CatBatchLookup`
  also contains a `Sub CatBatchLookup`, and in a compile-time reference the module
  name wins. `CatRibbon` calls it module-qualified as `CatBatchLookup.CatBatchLookup`.
- **"User-defined type not defined"** — the Office object library reference is
  missing; change `IRibbonControl` / `IRibbonUI` to `Object`, or add the reference.
- **Blank button icon** — an `imageMso` name does not exist in that Office
  build. It fails silently; nothing else breaks. Cosmetic.
- **A form button does nothing when clicked** — its `(Name)` does not match the
  handler, or the form's code is a version behind. Double-click the button in
  the designer: if it lands in an *empty* sub, the code needs re-pasting.
  `FitControls` reaches the optional buttons through `Me.Controls(...)` inside
  `On Error Resume Next`, which is what makes them optional — and also why a
  misnamed one is skipped without complaint.
- **"Ambiguous name detected"** after pasting form code — an empty stub VBA
  created when you double-clicked a new button is still there. Select all,
  delete, then paste.
- **A button stretching under its neighbour** — the form's code predates that
  button, so the layout is still dividing the row between the old ones.

## What this still doesn't fix

The add-in talks to the Azure proxy directly, so **none of the Snowflake
business rules apply** and nothing here writes a `CCAT_AUDIT` row. The write log
covers the local question — who changed what, and what it was before — but it
lives under one person's profile, not in the warehouse.

Showing NAXT and CCAT side by side *inside* the lookup window would need
Snowflake exposed outward: permission, an Azure Function to serve it, and a
service account so access is not tied to one person's role. **Compare to CCAT**
delivers most of that value without any of it, because you bring the comparison
data as a sheet.

**Cat cannot be asked what is pending for a dealer.**
`/ownershipRecords/search` takes at most two `stringEquals` filters over DCN /
asset name / serial number / make code — nothing on `ownershipRequestType`,
dealer or status. Not from here, and not from a new proxy route either. Finding
inbound transfer requests means sweeping a known list of DCNs and filtering the
results client-side.
