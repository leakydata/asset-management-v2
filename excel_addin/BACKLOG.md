# Cat Asset Tools — backlog

Everything we've agreed is worth building, in the order it's worth building it.
Tick items as they land. Add, drop and reorder freely — the point is that
nothing quietly falls off the end.

## How we work this list

**Batch the form changes.** Every change to a UserForm's code has to be pasted
into the VBE by hand — a `.frm` can't be imported as text the way a `.bas` can.
Doing that three times for three small changes is three chances to paste into
the wrong module. So items marked **[form]** get done in one pass, together, at
the point we're ready for all of them.

**`.bas` changes are cheap.** Re-import and go. Prefer putting logic in a
module and calling it from the form over writing it in the form.

**Anything that writes gets a dry run first.** Validate must be able to show
what a Run would do before Run does it. That principle is why §2 exists.

Legend: **[form]** touches a UserForm · **[bas]** module only ·
**[xml]** ribbon · **[?]** has an open question to resolve first

---

## 1. Ship-ready for other people

The installer means colleagues get this soon. These are the difference between
"it works" and a week of support calls.

- [x] **1.1 Test Connection button in Settings** **[bas]** **[xml]**
  Runs a known-good lookup and reports plainly: *no URL set* / *no key set* /
  *couldn't reach the proxy* / *key rejected (401/403)* / *worked — N records*.
  **Why:** the first support call will be "it doesn't work", and it'll be a
  typo'd function key. This answers it without you.
  **Done when:** each of the five states can be produced deliberately and each
  says something a non-technical user can act on.
  **Effort:** small. Best ratio on the whole list.
  **DONE.** Ribbon button in Setup, and offered automatically after Settings
  saves — a typo is cheapest to catch the second it is typed. Status codes were
  measured against the live proxy, not assumed: unknown serial 200, bad/absent
  key 401, wrong path 404, bad host no reply. The probe searches a serial that
  deliberately does not exist, so the test never breaks because someone expired
  the asset it was looking for.

- [x] **1.2 Plain-English API errors** **[bas]**
  Map the common proxy/CCAT failures to sentences. *"This DCN isn't ours —
  Cat only returns DCNs for B150 records"* rather than a raw 403 body.
  **Why:** a red Result cell full of JSON teaches the user nothing and sends
  them to you.
  **Done when:** 401, 403, 404, 429, 5xx and a timeout each have a mapped
  message, with the raw text kept as a fallback for anything unmapped.
  **Effort:** small.
  **DONE.** Rewrote `ProxyError` itself rather than adding a layer, so both
  call sites — `CatSearch` and `SendRow`'s Result cell — improved with no
  call-site changes. Fixed a real bug on the way: it read `error` as an object
  with code/description, but the proxy emits a flat string, so every
  proxy-generated error fell through to dumping raw JSON in the cell. Both
  shapes now parse. The distinction worth having is **401 vs 502** — 401 is
  your key and you can fix it; 502 means the proxy took your key fine and then
  couldn't reach or sign in to Cat, which no user can fix and none should be
  sent to Settings to try.

---

## Housekeeping — found along the way, not features

- [x] **H.2 Lookups are logged too, in their own file**
  Raised twice from use: doing two searches and then opening the log showed
  nothing, because the log covered writes only. That was correct by design and
  wrong for the person using it — a search is activity, and a log you check
  after doing work should show the work you did.
  **DONE.** `cat-lookup-log-YYYY-MM.csv`, deliberately **separate** from the
  write log: reads and writes answer different questions, and a few hundred
  searches would bury the handful of writes that matter. Logged inside
  `CatSearch`, the one place every window and every batch passes through, so
  neither form needed changing. Worksheet-function calls are skipped —
  `=CatLookupSerial()` re-evaluates on any recalc, and that is Excel, not a
  person. Failed lookups are logged before the raise, since "I searched and got
  a 401" is exactly the line you want afterwards.

- [x] **H.1 `azure-function/` in this repo is stale**
  It defines only the `search` route, but the add-in also calls `ownership`,
  `expire` and `transfer`. The deployed function is ahead of the source here.
  **Why it matters:** the repo is no longer the source of truth for the proxy,
  so the next person to change it will be working from an incomplete copy — and
  the error bodies for the three write routes can't be read from source.
  **Done when:** the folder matches what's deployed, or says plainly that it
  doesn't and points at where the real source lives.
  **DONE** — the second way. `azure-function/README.md` now opens with a
  warning that the folder is behind the deployment, names the three missing
  routes, and says **do not redeploy from it** (that would delete the write
  routes and break every operation sheet). Recovering the real source is a
  separate job for whoever next needs to change the proxy.

---

## 2. Make the destructive path safe

The two highest-value items here. Expire removes ownership records; today a
bad batch is unrecoverable except from memory.

- [x] **2.1 Validate shows the diff, not just the shape** **[bas]**
  Validate already checks required fields. Have it also fetch the current
  record and report what would actually change:
  ```
  Serial      Change
  KXH10658    OwnershipType  RENTAL → OWNED
  9303        no change
  MFG00123    NEW record
  ```
  **Why:** the README documents the footgun — *every update rewrites
  `ownershipTypeCode`, so pasting a stale value overwrites a correct one*.
  Nothing warns you today. This turns an invisible overwrite into a line you
  have to read past.
  **Cost:** one extra lookup per row on Validate. Acceptable — Validate is the
  slow, careful step by design.
  **Done when:** a sheet whose values match CCAT reports "no change" on every
  row, and a deliberately stale OwnershipType shows the arrow.
  **DONE.** Verified against live CCAT: a row pasted from a fresh lookup gives
  "no change", a stale one gives `CHANGES: OwnershipType RENTAL -> OWNED`.
  Expire names the record it would remove; Transfer says when there is no
  pending request to act on. Two rules it follows: it is **advisory only** — a
  diff never turns an OK row into SKIPPED, because a row's fate must not depend
  on a live call — and it **never fails the validation**, so Validate still
  works with the proxy down. Only fields the row will actually send are
  compared, since a blank cell is omitted from the request and can't change
  anything. Results cached per serial per run, and sheets over 50 rows ask
  first (one lookup per serial adds up).

- [x] **2.2 Snapshot every row before a write** **[bas]**
  Before Run sends anything, capture the current record for each row to a
  hidden sheet (or a CSV under `%APPDATA%`), keyed by run id + timestamp.
  **Why:** this is what makes 2.3 possible. Cheap on its own, transformative
  with the next item.
  **Done when:** a Run leaves behind a retrievable before-image of every row it
  touched, including rows that then failed.
  **DONE**, in `CatAudit.bas`, as one line per row shared with 2.4. Records
  three states, not two: FOUND / NONE / UNAVAILABLE. Undo must tell "there was
  no record" from "the lookup failed", or it writes blanks over a live record.
  Costs one lookup per distinct serial (shared with 2.1's cache) — always on,
  never prompted: a safety net you opt into is a safety net that is off.

- [x] **2.3 Rebuild an undo sheet from a snapshot** **[bas]** **[xml]**
  "Undo run &lt;id&gt;" produces a Cat Add-Update sheet that restores the before
  values. It does **not** auto-send — it hands you a sheet you Validate and Run
  yourself.
  **Why:** Expire is the operation that can ruin an afternoon. This is the
  feature that lets people use it without flinching.
  **Done when:** expiring a test record and then running the generated undo
  sheet restores it.
  **DONE.** **CCAT > Build Sheet > Undo a Run...** lists recent runs, and
  builds `Cat Undo <runid>` — its own sheet, never over `Cat Add-Update`, since
  someone recovering from a bad batch may have work in progress there. It
  **sends nothing**: an undo that sends is a second unreviewed write on top of
  the first, at the moment the person is rattled and least likely to check.
  Three cases handled honestly — FOUND restores, NONE (the run *created* the
  record) is offered as a separate Expire sheet behind its own confirmation
  because reversing it is destructive, and UNAVAILABLE is reported and never
  restored as blanks. Rows that failed at the time are skipped: they changed
  nothing, so restoring them would push a stale value over whatever is there
  now. Round-trip verified on live data — 22-field before-image through the CSV
  and back into the right sheet columns.

**§2 complete.** The destructive path now warns before, records during, and can
be reversed after.

- [x] **2.4 Write log** **[bas]**
  Timestamp, user, operation, request, response, result — appended for every
  write. Hidden sheet or `%APPDATA%` file.
  **Why:** the README names this gap outright: *"nothing here writes a
  CCAT_AUDIT row… manual edits are the ones most worth having a record of."*
  When someone asks who expired a record on Tuesday, there's an answer.
  **Done when:** every Run appends one line per row, and the log survives
  closing the workbook.
  **Note:** overlaps 2.2 — likely one storage mechanism serving both.
  **DONE.** One mechanism, as predicted. `%APPDATA%\CatAssetTools\cat-write-log-YYYY-MM.csv` — outside the workbook, so it outlives the file
  that happened to be open. **CCAT > Write Log** opens the folder, and the run
  summary quotes the run id. Every field is quoted: a DCN name like "SMS RENTAL
  (WA) PTY LTD, INC" would shift every later column, and a naive `Split(",")`
  gave 17 fields where 15 were written — verified by round-trip test.

---

## 3. Run robustness

- [x] **3.1 Esc cancels a Run, cleanly** **[bas]**
  Stops between rows, never mid-request. Rows already sent keep their result.
  **Done when:** cancelling a 100-row run leaves rows 1..n green and the rest
  untouched, with a clear "cancelled at row n" message.
  **DONE**, via `GetAsyncKeyState` rather than `Application.EnableCancelKey`.
  EnableCancelKey can interrupt mid-HTTP-call, which on a write means the
  request may have gone through with nothing recorded — and disabling it around
  the send just swallows the keypress instead. Reading the key state directly
  has neither problem: a press *during* a request is still waiting to be found
  at the top of the next row, so the stop happens exactly where we choose. A
  stale press from before the run is drained first. The summary leads with
  "STOPPED at row n", because "17 rows processed" is not the headline when 483
  were never sent.

- [x] **3.2 Run only rows without a green result** **[bas]**
  Resume a partially-completed batch instead of re-sending everything.
  **Why:** Add/Update is idempotent so re-running is *probably* harmless, but
  Expire isn't obviously so, and "just run it all again" is an uncomfortable
  instruction on a write path.
  **Done when:** a run interrupted at row 300 of 500 can be finished with one
  click and no duplicate sends.
  **DONE**, and with **no new ribbon button** — dropped the [xml] tag. Run
  detects rows that already came back OK and asks whether to skip them, so the
  question appears exactly when it is relevant and never on a fresh sheet.
  Matches on the Result **text**, not the cell colour: colour is presentation,
  and re-formatting a sheet must not change what gets re-sent.

- [x] **3.3 Progress in the status bar** **[bas]**
  Row n of N, and the serial currently in flight.
  **ALREADY DONE — do not build twice.** `RunCore` has always set
  `Application.StatusBar = OpLabel(op) & " " & done & " of " & nRows & " (" &
  serial & ")"`. Found while implementing 2.1. Listing it as a goal was my
  error, not a gap.

---

## 4. Daily-use delight

All of the **[form]** work lives here. Do it in one sitting.

- [x] **4.1 Recent lookups dropdown** **[form]** **[bas]**
  Replace `txtSerial` with a **ComboBox** — type a value as now, or pick from
  the last 25 searched.
  **The detail that makes this cheap:** an MSForms ComboBox supports `.Text`,
  `.SetFocus`, `.SelStart` and `.SelLength` — every member the current code
  uses. Keep the control **named `txtSerial`** and swap only its type, and no
  existing line of form code changes. `Style` stays `fmStyleDropDownCombo` (the
  default) so free typing still works.
  **Storage:** HKCU alongside the proxy settings, one delimited string. Cap in
  a constant so "25 or more" is one number.
  **Two rules worth having:**
  - separate lists for serials and DCNs — never mixed
  - only remember a lookup that **returned something**, so typos and dead
    serials don't silt up the list
  **Done when:** looking up five serials, closing Excel, reopening and clicking
  the arrow shows those five, newest first, no duplicates.
  **DONE.** Storage and loading live in `CatLookShared`, so the form gained two
  lines. `MruLoad` takes the control as `Object` and swallows errors, so it does
  nothing at all to a plain TextBox — the code is safe whether or not the
  control has been swapped yet.

- [x] **4.2 "Send this list to a sheet"** **[form]** **[bas]**
  A button on the DCN window that writes the current result list to a results
  sheet.
  **Why:** you look up a DCN, see 55 assets, and today you have to go run Batch
  DCNs separately to get them into a sheet. The data is already in `mRows`.
  **Done when:** the sheet it writes is identical to what Batch DCNs produces
  for that DCN.
  **DONE** — same *shape* rather than byte-identical (`QueryDCN` + the 22
  fields + `Note`), which matters for more than tidiness: Validate and Run
  refuse any sheet carrying a QuerySerial/QueryDCN column, so matching the
  shape inherits that refusal and a results sheet can't be fired off as an
  operation sheet.

- [x] **4.4 Say "not found" where the eye actually is** **[form]**
  When a lookup returns zero records, both floating windows say so **only on
  the title bar** while the two lists go blank. Blank lists read as "it broke",
  not "the answer is none" — and the title bar is the smallest text on the
  form.
  **The other two surfaces are already fine**, which is why this is narrow:
  `=CatLookupSerial()` returns *"No records found for X"*, and batch lookup
  writes `NOT IN CCAT` / `NO ACTIVE RECORDS` per row plus a summary count. Only
  the windows are quiet.
  **Fix:** put one row into the record list itself — *"No records — KR12345 is
  not in CCAT"* — and keep the caption as well. No new control, so it is pure
  code, and `ShowRecord` already guards on `mRows Is Nothing` so a click on
  that row does nothing.
  **Worth distinguishing**, the way batch lookup already does: a serial that
  isn't in CCAT at all, versus a DCN that exists but currently owns nothing.
  **Done when:** looking up a nonsense serial puts a visible line in the list,
  and it clears on the next successful search.
  **Batches with 4.1 and 4.2** — all three are form-code, so they get pasted
  into the VBE once, not three times.
  **DONE.** One line into the record list on both windows, alongside the
  caption.

**§4 complete.** One VBE paste covered all of it, which is what the batching
rule was for.

- [x] **4.3 Right-click a cell ▸ "Cat: look up this serial"** **[bas]**
  Context-menu entry, seeded from the clicked cell.
  **Why:** beats travelling to the ribbon for the most common action.
  **Note:** needs the menu torn down on add-in unload or it lingers in Excel
  after the add-in is gone.
  **DONE, and it needed no form change at all.** Installed from
  `CatRibbonOnLoad`, which fires when the .xlam loads and is the only startup
  hook shippable in a `.bas`. `Temporary:=True` answers the teardown note —
  Excel drops the controls on close, so an uninstalled add-in can't leave an
  entry pointing at a missing macro. A tagged sweep before adding handles the
  other case, a reload inside one session stacking duplicates.

---

## 5. Bigger, later

- [x] **5.0 Reconcile a NAXT sheet against CCAT**  ← the actual job **[bas]** **[xml]**
  Point it at a sheet of NAXT values. It looks each serial up in CCAT and
  writes back **only the rows that disagree**, pre-filled with the NAXT values
  in Add/Update shape, ready to Validate and Run. Rows that already match are
  dropped rather than reviewed. Serials CCAT has never heard of come out as
  ADD rows.
  **Why this is now first:** the add-in exists for people who check serials and
  either add them to CCAT or correct CCAT to match NAXT. Nothing built so far
  does the *comparing* — 2.1 tells you a row differs after you have decided to
  send it; this decides which rows are worth sending at all. It collapses
  export → Batch Serials → eyeball two sets of columns → build a sheet → paste
  → Validate → Run into two steps.
  **Needs no new infrastructure**, which is the point: the NAXT data arrives as
  a sheet the user already has open, so there is no Snowflake connection, no
  new Azure Function, no service account and no permission to seek.
  **Mostly parts we already have:** the batch sweep, the field comparison from
  2.1, the sheet writer from 4.2, and header matching that already ignores
  case, spaces and punctuation.
  **DONE.** **CCAT > Compare to CCAT**, acting on the active sheet. Reads
  whatever recognised columns are there, compares only cells that are filled
  in (a blank is omitted from the request, so calling it a difference would be
  a lie), and writes only the rows that disagree.
  **The DCN question answered itself in testing, and the answer matters:**
  both test serials sit on **two** ownership records, so without a DCN column
  almost every row is ambiguous — and dealer code doesn't rescue it, since both
  records are B150 too. That is the data, not the code. So a NAXT export
  **wants a DCN column**; without one the feature only handles single-record
  serials and brand-new ones.
  Ambiguous rows go **onto the sheet** with the DCN blank rather than into a
  footnote — Validate then marks them SKIPPED with a reason, turning a dead end
  into a list of rows needing a DCN. A number in a dialog is not actionable; a
  yellow row is.

- [ ] **5.2 Snowflake reconciliation in the lookup window** **[?]** **[form]**
  Show *"CCAT says OWNED, our records say RENTAL"* beside the record, from the
  `STD_UMT` tracking tables and the discrepancy sweep.
  **BLOCKED, and not on engineering.** Confirmed 2026-09-03: it needs
  permission to expose Snowflake data outward, a new Azure Function to serve
  it, and a service account so the access is not tied to one person's role.
  That is an organisational sequence, not a sprint item.
  **5.0 delivers most of the same value without any of it**, because the user
  brings the comparison data themselves. Revisit only if the reconcile proves
  the round trip through a spreadsheet is the bottleneck.

---

## Dropped

- **5.1 "Waiting on me" — pending transfers.** `OwnershipRequestType =
  RECEIVED` means another dealer wants ownership moved, and approving those is
  a different role from reconciling NAXT. A button these users never press is
  worse than no button. Bring it back if the same people do handle transfers.
  The research is not wasted and is worth keeping: **Cat cannot be asked "what
  is pending for us" at all** — `/ownershipRecords/search` takes at most two
  `stringEquals` filters over DCN / asset name / serial number / make code, and
  offers nothing on `ownershipRequestType`, dealer or status. Not from the
  add-in, and not from a new proxy route either. Any future version of this
  must sweep a known list and filter client-side.

---

## Resolved / not doing

- **Full shared body between the two lookup forms.** Considered and declined —
  shared code would have to take the form as a late-bound `Object`, turning
  control names the compiler checks today into strings that fail at runtime,
  and it would mean re-pasting both forms by hand. Only the paste rules are
  shared, in `CatLookShared`. Revisit if a third window ever appears.
