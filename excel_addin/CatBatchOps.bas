Attribute VB_Name = "CatBatchOps"
'==============================================================================
' Cat Asset Management V2 - sheet-driven operations (via PROXY)
'
' Replaces BOTH the old CatActions (a form of named cells) and CatBatchActions.
' Every operation is now the same shape: a sheet of rows with a Result column.
' A single asset is just a batch of one - type or paste one row and run it.
'
'   Cat Add-Update   POST /api/ownership   add or update ownership records
'   Cat Expire       POST /api/expire      remove ownership records
'   Cat Transfer     POST /api/transfer    approve / reject pending transfers
'
' COLUMN ORDER IS THE INTERFACE. The first columns of each sheet are exactly
' the leading columns of the batch-lookup output, in the same order, so moving
' data across is a drag-select and a paste - no Ctrl-clicking:
'
'   lookup output :  1 SerialNumber  2 MakeCode  3 DCN  4 OwnershipType
'                    5 Model         6 ModelYear    ... then reference detail
'
'   Cat Transfer   <- lookup columns 1-2   (Serial, Make Code)
'   Cat Expire     <- lookup columns 1-3   (+ DCN)
'   Cat Add-Update <- lookup columns 1-6   (+ Ownership Type, Model, Model Year)
'
' Each action's required set is a prefix of the next, so one lookup ordering
' serves all three. On a results sheet the fields start in column B (column A
' is QuerySerial / QueryDCN), so the drags are B:C, B:D and B:G.
'
' Dealer Make Code is deliberately LAST on every sheet. It is the alternative
' to Make Code - supplying both is rejected - so it must not sit inside the
' paste block. Everything else after the block is optional.
'
' Keep HeaderArray (CatAssetLookup) and SheetSpec (below) in step. Reordering
' one without the other silently breaks the paste.
'
' Headers are matched by NAME (case, spaces and punctuation ignored), so extra
' columns and reordering are both harmless.
'
' TWO BUTTONS DRIVE ALL THREE. Validate and Run look at the ACTIVE sheet and
' work out the operation from its name, falling back to its headers. Validate
' sends nothing - it checks every row and reports what would happen.
'
' Depends on CatAssetLookup (CatAddUpdate, CatExpire, CatTransfer, CleanId,
' FieldOf, ProxyError, TargetBook) and VBA-JSON.
'
' WARNING: Run changes real data in CCAT. There is no Cat sandbox. Always
' Validate first - that is the guardrail, especially on Expire, where a bad
' paste is a lot of ownership records.
'==============================================================================
Option Explicit

Public Const SH_ADD As String = "Cat Add-Update"
Public Const SH_EXP As String = "Cat Expire"
Public Const SH_TRF As String = "Cat Transfer"

Private Const OP_ADD As String = "ADD_UPDATE"
Private Const OP_EXP As String = "EXPIRE"
Private Const OP_TRF As String = "TRANSFER"

Private Const HEADER_ROW As Long = 1
Private Const MAX_ATTEMPTS As Long = 3
Private Const RETRY_WAIT_SECONDS As Long = 1

' Above this many rows, Validate asks before comparing against CCAT - the
' comparison costs one lookup per distinct serial, and on a 500-row sheet that
' is minutes of waiting nobody asked for.
Private Const DIFF_PROMPT_ROWS As Long = 50

' serial -> Collection of RecordValues arrays, or Empty when that serial's
' lookup failed. Per-run, so a sheet with the same serial on ten rows costs one
' call, and a serial that errors is not retried ten times.
Private mDiffCache As Object

'==============================================================================
' PUBLIC: the two buttons
'==============================================================================

' Checks every row and reports what WOULD happen. Sends nothing.
Public Sub CatValidateSheet()
    RunCore True
End Sub

' Sends every valid row. Confirms once, with the count.
Public Sub CatRunSheet()
    RunCore False
End Sub

'==============================================================================
' PUBLIC: undo a run
'
' Rebuilds what a past Run overwrote, from the before-images in the write log.
' This is the reason the log captures them.
'
' IT BUILDS A SHEET AND STOPS. It does not send. An undo that sends is a second
' unreviewed write on top of the first, at the moment someone is rattled and
' least likely to check - which is how one bad batch becomes two. You get a
' sheet, you Validate it, you Run it, exactly like any other work.
'
' WHAT IT CAN AND CANNOT PUT BACK:
'
'   BeforeState  What happened            Undo
'   FOUND        a record was changed     restore the old values (Add/Update)
'                or expired
'   NONE         a record was CREATED     the reverse is an Expire, offered
'                                         separately and confirmed on its own
'   UNAVAILABLE  the lookup failed first  nothing - we never knew the old value
'
' UNAVAILABLE rows are reported, never silently dropped and never restored as
' blanks. Writing emptiness over a live record because a lookup timed out is
' exactly the accident this whole section exists to prevent.
'
' Failed rows are skipped too: a row that did not go through changed nothing,
' so "restoring" it would push a stale value over whatever is there now.
'==============================================================================
Public Sub CatUndoRun()
    On Error GoTo Fail

    Dim runs As Collection: Set runs = AuditRunIds()
    If runs.Count = 0 Then
        MsgBox "No runs have been logged this month." & vbCrLf & vbCrLf & _
               "Undo reads the write log, which gets a line for every row Run " & _
               "sends. Nothing has been sent yet, so there is nothing to undo.", _
               vbInformation, "Cat Asset Tools - Undo"
        Exit Sub
    End If

    Dim runId As String
    runId = PickRun(runs)
    If Len(runId) = 0 Then Exit Sub

    Dim lines As Collection: Set lines = AuditLinesFor(runId)
    If lines.Count = 0 Then
        MsgBox "No logged rows found for run " & runId & ".", vbExclamation, _
               "Cat Asset Tools - Undo"
        Exit Sub
    End If

    ' Sort the run's rows into what can be put back, what would need expiring,
    ' and what we simply do not know.
    Dim restorable As Collection: Set restorable = New Collection
    Dim created As Collection:    Set created = New Collection
    Dim unknown As Long, failed As Long

    Dim i As Long, f As Variant
    For i = 1 To lines.Count
        f = lines(i)
        If UBound(f) >= 32 Then
            If Not OutcomeWasOk(CStr(f(9))) Then
                failed = failed + 1
            Else
                Select Case CStr(f(10))
                    Case BEFORE_FOUND:       restorable.Add f
                    Case BEFORE_NONE:        created.Add f
                    Case BEFORE_UNAVAILABLE: unknown = unknown + 1
                End Select
            End If
        End If
    Next i

    If restorable.Count = 0 And created.Count = 0 Then
        MsgBox "Nothing in run " & runId & " can be undone." & vbCrLf & vbCrLf & _
               failed & " row(s) failed at the time, so they changed nothing." & vbCrLf & _
               unknown & " row(s) had no before-image (the lookup failed first), " & _
               "so the old values were never known.", _
               vbExclamation, "Cat Asset Tools - Undo"
        Exit Sub
    End If

    Dim built As String

    If restorable.Count > 0 Then
        built = BuildUndoSheet(runId, restorable)
    End If

    ' Records this run CREATED. Reversing those means expiring them, which is
    ' destructive in its own right - so it is a separate question, asked
    ' plainly, never bundled into "undo".
    If created.Count > 0 Then
        If MsgBox(created.Count & " row(s) in this run CREATED a record that did " & _
                  "not exist before." & vbCrLf & vbCrLf & _
                  "Undoing those means EXPIRING them - removing the ownership " & _
                  "records this run added. That is a destructive operation in " & _
                  "its own right, so it is a separate sheet and a separate " & _
                  "decision." & vbCrLf & vbCrLf & _
                  "Build an Expire sheet for them?", _
                  vbExclamation + vbYesNo + vbDefaultButton2, _
                  "Cat Asset Tools - Undo") = vbYes Then
            built = built & IIf(Len(built) > 0, " and ", "") & _
                    BuildUndoExpireSheet(runId, created)
        End If
    End If

    If Len(built) = 0 Then Exit Sub

    MsgBox "Built " & built & " from run " & runId & "." & vbCrLf & vbCrLf & _
           "NOTHING HAS BEEN SENT. Check the sheet, then Validate, then Run - " & _
           "the same as any other work." & vbCrLf & vbCrLf & _
           IIf(unknown > 0, unknown & " row(s) left out: no before-image, so the " & _
               "old values were never known." & vbCrLf, "") & _
           IIf(failed > 0, failed & " row(s) left out: they failed at the time, " & _
               "so they changed nothing." & vbCrLf, ""), _
           vbInformation, "Cat Asset Tools - Undo"
    Exit Sub

Fail:
    MsgBox "Error: " & Err.Description, vbCritical, "Cat Asset Tools - Undo"
End Sub

' SendRow returns "OK (200) ..." on success and "FAILED..." otherwise.
Private Function OutcomeWasOk(ByVal outcome As String) As Boolean
    OutcomeWasOk = (Left$(LTrim$(outcome), 2) = "OK")
End Function

' A numbered list in an InputBox. Not a UserForm on purpose: this is the one
' module that must keep working when someone is mid-panic, and a .bas needs no
' hand-pasting into the VBE to arrive.
Private Function PickRun(ByVal runs As Collection) As String
    Const SHOW As Long = 15

    Dim prompt As String, i As Long, r As Variant
    prompt = "Which run do you want to undo?" & vbCrLf & vbCrLf
    For i = 1 To runs.Count
        If i > SHOW Then Exit For
        r = runs(i)
        prompt = prompt & i & ".  " & r(1) & "   " & OpLabel(CStr(r(2))) & _
                 "   on " & r(3) & vbCrLf
    Next i
    prompt = prompt & vbCrLf & "Type a number (1-" & _
             IIf(runs.Count < SHOW, runs.Count, SHOW) & "):"

    Dim answer As String
    answer = InputBox(prompt, "Cat Asset Tools - Undo a Run")
    If Len(Trim$(answer)) = 0 Then Exit Function

    Dim n As Long
    On Error Resume Next
    n = CLng(Trim$(answer))
    On Error GoTo 0
    If n < 1 Or n > runs.Count Or n > SHOW Then
        MsgBox "That is not one of the numbers listed.", vbExclamation, _
               "Cat Asset Tools - Undo"
        Exit Function
    End If

    r = runs(n)
    PickRun = CStr(r(0))
End Function

' The restore sheet: an Add/Update carrying the values as they were.
'
' Written to its OWN sheet, never over 'Cat Add-Update'. Someone recovering
' from a bad batch may well have work in progress there, and silently clearing
' it while they are already having a bad afternoon is not on.
Private Function BuildUndoSheet(ByVal runId As String, ByVal rows As Collection) As String
    Dim ws As Worksheet
    Set ws = FreshSheet("Cat Undo " & runId)

    Dim nm As String, headers As Variant, notes As Variant, tiers As Variant
    SheetSpec OP_ADD, nm, headers, notes, tiers
    WriteHeaders ws, headers, tiers

    ' Sheet column <- before-image index. CustomAssetName is deliberately
    ' absent: the API returns assetName as a coalesced custom-or-base value, so
    ' there is nothing to restore it from honestly, and a blank cell is omitted
    ' from the request rather than overwriting anything. Dealer Make Code is
    ' left blank too - Make Code is being supplied, and the API rejects both.
    Dim keys As Variant, idx As Variant
    keys = Array("serial", "makecode", "dcn", "ownershiptype", "model", _
                 "modelyear", "productfamilycode", "productfamilyname", "baseassetname")
    idx = Array(0, 1, 2, 3, 4, 5, 18, 19, 21)

    Dim cols As Object: Set cols = HeaderMap(ws)
    Dim i As Long, k As Long, c As Long, f As Variant
    For i = 1 To rows.Count
        f = rows(i)
        For k = LBound(keys) To UBound(keys)
            c = ColOf(cols, CStr(keys(k)))
            If c > 0 Then
                ws.Cells(HEADER_ROW + i, c).NumberFormat = "@"
                ws.Cells(HEADER_ROW + i, c).Value = CStr(f(11 + idx(k)))
            End If
        Next k
    Next i

    ws.Columns.AutoFit
    ws.Activate
    BuildUndoSheet = "'" & ws.Name & "' (" & rows.Count & " row(s) to restore)"
End Function

' The reverse of an Add: expire what the run created.
Private Function BuildUndoExpireSheet(ByVal runId As String, ByVal rows As Collection) As String
    Dim ws As Worksheet
    Set ws = FreshSheet("Cat Undo Expire " & runId)

    Dim nm As String, headers As Variant, notes As Variant, tiers As Variant
    SheetSpec OP_EXP, nm, headers, notes, tiers
    WriteHeaders ws, headers, tiers

    ' Expire needs Serial, Make Code and DCN, and those come from the LOG LINE
    ' rather than the before-image - there was no before-image, that is the
    ' whole point of this sheet.
    Dim cols As Object: Set cols = HeaderMap(ws)
    Dim cS As Long: cS = ColOf(cols, "serial", "serialnumber")
    Dim cM As Long: cM = ColOf(cols, "makecode")
    Dim cD As Long: cD = ColOf(cols, "dcn")

    Dim i As Long, f As Variant
    For i = 1 To rows.Count
        f = rows(i)
        If cS > 0 Then ws.Cells(HEADER_ROW + i, cS).NumberFormat = "@": ws.Cells(HEADER_ROW + i, cS).Value = CStr(f(7))
        If cD > 0 Then ws.Cells(HEADER_ROW + i, cD).NumberFormat = "@": ws.Cells(HEADER_ROW + i, cD).Value = CStr(f(8))
        ' Make Code was not logged separately; it is on the before-image only
        ' when there was one. Left blank here for you to fill - Validate will
        ' say so rather than letting it through.
        If cM > 0 Then ws.Cells(HEADER_ROW + i, cM).NumberFormat = "@"
    Next i

    ws.Columns.AutoFit
    BuildUndoExpireSheet = "'" & ws.Name & "' (" & rows.Count & " row(s) to expire - " & _
                           "Make Code needs filling in)"
End Function

' A sheet with this name, emptied if it already exists. Undo sheets are named
' after the run, so re-undoing the same run reuses its sheet rather than
' littering the workbook with near-identical tabs.
Private Function FreshSheet(ByVal nm As String) As Worksheet
    Dim wb As Workbook: Set wb = TargetBook()
    Dim ws As Worksheet, shp As Shape

    ' Excel sheet names cap at 31 characters.
    If Len(nm) > 31 Then nm = Left$(nm, 31)

    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = nm
    Else
        ws.Cells.Clear
        For Each shp In ws.Shapes: shp.Delete: Next shp
    End If

    Set FreshSheet = ws
End Function

Private Sub WriteHeaders(ByVal ws As Worksheet, ByVal headers As Variant, ByVal tiers As Variant)
    Dim c As Long, bg As Long, fg As Long
    For c = 0 To UBound(headers)
        Select Case tiers(c)
            Case 0: bg = RGB(31, 78, 121): fg = vbWhite
            Case 1: bg = RGB(46, 117, 182): fg = vbWhite
            Case 3: bg = RGB(89, 89, 89): fg = vbWhite
            Case Else: bg = RGB(189, 215, 238): fg = RGB(31, 31, 31)
        End Select
        With ws.Cells(HEADER_ROW, c + 1)
            .Value = headers(c)
            .Font.Bold = True
            .Font.Color = fg
            .Interior.Color = bg
        End With
    Next c
End Sub

'==============================================================================
' PUBLIC: build the input sheets
'==============================================================================
Public Sub CatBuildAddUpdateSheet()
    BuildSheet OP_ADD
End Sub

Public Sub CatBuildExpireSheet()
    BuildSheet OP_EXP
End Sub

Public Sub CatBuildTransferSheet()
    BuildSheet OP_TRF
End Sub

'==============================================================================
' PRIVATE: the engine
'==============================================================================
Private Sub RunCore(ByVal dryRun As Boolean)
    On Error GoTo Fail

    Dim wb As Workbook: Set wb = TargetBook()
    Dim ws As Worksheet: Set ws = ActiveSheet
    If ws Is Nothing Then Err.Raise vbObjectError + 30, , "No active sheet."

    Dim op As String, byName As Boolean
    op = OperationOf(ws, byName)
    If Len(op) = 0 Then
        MsgBox "This sheet isn't one of the Cat operation sheets." & vbCrLf & vbCrLf & _
               "Use Build Sheet to create '" & SH_ADD & "', '" & SH_EXP & "' or '" & _
               SH_TRF & "', or switch to one you already have.", _
               vbExclamation, "Cat Asset Tools"
        Exit Sub
    End If

    ' A renamed or hand-built sheet was matched on its headers alone. Say which
    ' operation that resolved to before doing anything with it.
    If Not byName Then
        If MsgBox("'" & ws.Name & "' isn't a standard Cat sheet name." & vbCrLf & vbCrLf & _
                  "From its columns this looks like a " & OpLabel(op) & " sheet." & vbCrLf & _
                  "Treat it as " & OpLabel(op) & "?", _
                  vbQuestion + vbYesNo + vbDefaultButton2, "Cat Asset Tools") <> vbYes Then Exit Sub
    End If

    Dim cols As Object: Set cols = HeaderMap(ws)
    Dim cSerial As Long: cSerial = ColOf(cols, "serial", "serialnumber")
    If cSerial = 0 Then
        MsgBox "No 'Serial' column in row " & HEADER_ROW & " of '" & ws.Name & "'.", _
               vbExclamation, "Cat Asset Tools"
        Exit Sub
    End If

    Dim cResult As Long: cResult = EnsureResultColumn(ws, cols)

    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, cSerial).End(xlUp).Row
    Dim nRows As Long: nRows = CountRows(ws, cSerial, lastRow)
    If nRows = 0 Then
        MsgBox "No serial numbers found under the headers.", vbExclamation, "Cat Asset Tools"
        Exit Sub
    End If

    If Not dryRun Then
        If Not ConfirmRun(op, nRows) Then Exit Sub
    End If

    ' Validate compares each row against what CCAT holds right now.
    Dim compare As Boolean
    If dryRun Then compare = AskCompare(nRows)

    ' Both paths use the per-serial cache: Validate to diff, Run to capture the
    ' before-image it logs.
    ResetDiffCache

    ' A Run is logged. Not optional and not prompted for - the run you would
    ' most want a record of is the one nobody expected to go wrong, and a
    ' safety net you are asked to opt into is a safety net that is off.
    Dim runId As String
    If Not dryRun Then runId = AuditBegin()

    Dim nOk As Long, nFailed As Long, nSkipped As Long, done As Long
    Application.ScreenUpdating = False

    Dim r As Long
    For r = HEADER_ROW + 1 To lastRow
        Dim serial As String: serial = CleanId(CStr(ws.Cells(r, cSerial).Value))
        If Len(serial) = 0 Then GoTo NextRow

        done = done + 1
        Application.StatusBar = OpLabel(op) & " " & done & " of " & nRows & "  (" & serial & ")"

        Dim why As String
        why = ValidateRow(ws, r, cols, op, serial)
        If Len(why) > 0 Then
            WriteResult ws, r, cResult, "SKIPPED: " & why, 2
            nSkipped = nSkipped + 1
            GoTo NextRow
        End If

        If dryRun Then
            ' Built with If, NOT IIf. IIf evaluates BOTH arms, so an IIf here
            ' would call DiffNote - and hit the API - even with comparison
            ' turned off.
            Dim note As String
            note = "OK to send" & PreviewNote(ws, r, cols, op)
            If compare Then note = note & DiffNote(ws, r, cols, op, serial)
            WriteResult ws, r, cResult, note, 1
            nOk = nOk + 1
            GoTo NextRow
        End If

        ' Capture what the record looks like BEFORE the send. Has to happen
        ' here: after SendRow the old values are gone, and on an Expire the
        ' whole record is.
        Dim beforeState As String, before As Variant
        beforeState = CaptureBefore(ws, r, cols, serial, before)

        Dim outcome As String, kind As Long
        outcome = SendRow(ws, r, cols, op, serial, kind)
        WriteResult ws, r, cResult, outcome, kind
        If kind = 1 Then nOk = nOk + 1 Else nFailed = nFailed + 1

        AuditRow runId, ws, r, op, serial, _
                 CleanId(CellStr(ws, r, ColOf(cols, "dcn"))), _
                 outcome, beforeState, before
NextRow:
    Next r

    ws.Columns.AutoFit
    Application.StatusBar = False
    Application.ScreenUpdating = True

    ' The run id is quoted here because it is what Undo asks for. Telling
    ' someone afterwards where the record of what they just did lives is the
    ' difference between having a log and having a log anyone uses.
    Dim tail As String
    If Not dryRun Then
        tail = vbCrLf & vbCrLf & "Logged as run " & runId & vbCrLf & _
               "(CCAT > Write Log, or Undo Run to put these back)"
    End If

    MsgBox IIf(dryRun, "VALIDATION ONLY - nothing was sent." & vbCrLf & vbCrLf, "") & _
           done & " row(s) " & IIf(dryRun, "checked", "processed") & ":" & vbCrLf & _
           "  - " & nOk & IIf(dryRun, " ready to send (green)", " OK (green)") & vbCrLf & _
           IIf(dryRun, "", "  - " & nFailed & " failed (red)" & vbCrLf) & _
           "  - " & nSkipped & " skipped (yellow - missing or invalid fields)" & tail, _
           IIf(nFailed > 0 Or nSkipped > 0, vbExclamation, vbInformation), _
           OpLabel(op) & IIf(dryRun, " - Validate", "")
    Exit Sub
Fail:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical, "Cat Asset Tools"
End Sub

'==============================================================================
' PRIVATE: which operation is this sheet?
'==============================================================================
Private Function OperationOf(ByVal ws As Worksheet, ByRef byName As Boolean) As String
    byName = True
    Select Case ws.Name
        Case SH_ADD: OperationOf = OP_ADD: Exit Function
        Case SH_EXP: OperationOf = OP_EXP: Exit Function
        Case SH_TRF: OperationOf = OP_TRF: Exit Function
    End Select

    byName = False
    Dim d As Object: Set d = HeaderMap(ws)

    ' A batch-LOOKUP results sheet also carries an OwnershipType column, so it
    ' would otherwise look like an Add-Update sheet and running it would fire
    ' writes for every row. Its Query* column is the tell - refuse outright.
    If ColOf(d, "queryserial", "querydcn") > 0 Then Exit Function

    If ColOf(d, "serial", "serialnumber") = 0 Then Exit Function

    If ColOf(d, "status") > 0 And ColOf(d, "reason") > 0 Then
        OperationOf = OP_TRF
    ElseIf ColOf(d, "ownershiptype", "ownershiptypecode", "type") > 0 Then
        OperationOf = OP_ADD
    ElseIf ColOf(d, "dcn") > 0 Then
        OperationOf = OP_EXP
    End If
End Function

Private Function OpLabel(ByVal op As String) As String
    Select Case op
        Case OP_ADD: OpLabel = "Add / Update"
        Case OP_EXP: OpLabel = "Expire"
        Case OP_TRF: OpLabel = "Transfer"
        Case Else:   OpLabel = op
    End Select
End Function

Private Function ConfirmRun(ByVal op As String, ByVal nRows As Long) As Boolean
    Dim msg As String
    msg = nRows & " row(s) will be sent as " & UCase$(OpLabel(op)) & " requests." & vbCrLf & vbCrLf

    Select Case op
        Case OP_EXP
            msg = msg & "EXPIRE REMOVES OWNERSHIP RECORDS." & vbCrLf & _
                  "Expiring a PENDING record also cancels its transfer request." & vbCrLf & vbCrLf & _
                  "If you have not run Validate on this sheet, stop and do that first." & vbCrLf & vbCrLf
        Case OP_ADD
            msg = msg & "Add / update is idempotent, so a failed row is safe to re-run." & vbCrLf & _
                  "Blank optional cells are left out of the request entirely and will" & vbCrLf & _
                  "NOT overwrite what CCAT already holds." & vbCrLf & vbCrLf
        Case OP_TRF
            msg = msg & "This approves or rejects pending transfer requests." & vbCrLf & vbCrLf
    End Select

    msg = msg & "This changes real data in CCAT. Continue?"
    ConfirmRun = (MsgBox(msg, vbExclamation + vbYesNo + vbDefaultButton2, _
                         "Confirm " & OpLabel(op)) = vbYes)
End Function

'==============================================================================
' PRIVATE: per-row validation. Returns "" when the row is good to send.
'==============================================================================
Private Function ValidateRow(ByVal ws As Worksheet, ByVal r As Long, ByVal cols As Object, _
                             ByVal op As String, ByVal serial As String) As String
    Dim mk As String: mk = CellStr(ws, r, ColOf(cols, "makecode", "make"))
    Dim dmk As String: dmk = CellStr(ws, r, ColOf(cols, "dealermakecode", "dealermake"))
    Dim dcn As String: dcn = CleanId(CellStr(ws, r, ColOf(cols, "dcn")))

    ' "One of makeCode or dealerMakeCode must be provided" - the API rejects
    ' both together and rejects neither. They are separate code systems for the
    ' same manufacturer (Cat says CAT, our dealer code is AA), so Make Code is
    ' the normal one and Dealer Make Code is only a fallback.
    If Len(mk) = 0 And Len(dmk) = 0 Then
        ValidateRow = "Make Code is required (or Dealer Make Code instead).": Exit Function
    End If
    If Len(mk) > 0 And Len(dmk) > 0 Then
        ValidateRow = "Make Code and Dealer Make Code are alternatives - clear one. " & _
                      "Use Make Code unless your data only has the dealer code.": Exit Function
    End If

    Select Case op
        Case OP_ADD
            If Len(dcn) = 0 Then ValidateRow = "DCN is required.": Exit Function

            ' Required on every row by choice, not by the API - the API only
            ' demands it for a NEW record. Catching it here turns the single
            ' most common 400 into a skipped row you can see and fix.
            Dim ot As String
            ot = LCase$(CellStr(ws, r, ColOf(cols, "ownershiptype", "ownershiptypecode", "type")))
            If Len(ot) = 0 Then
                ValidateRow = "Ownership Type is required.": Exit Function
            End If
            If InStr(1, "|owned|rental|leased|sold|inventory|unknown|", "|" & ot & "|") = 0 Then
                ValidateRow = "Ownership Type '" & ot & "' is not one of " & _
                              "owned/rental/leased/sold/inventory/unknown.": Exit Function
            End If

            Dim yr As String: yr = CellStr(ws, r, ColOf(cols, "modelyear", "year"))
            If Len(yr) > 0 Then
                If Len(yr) <> 4 Or Not IsNumeric(yr) Then
                    ValidateRow = "Model Year must be 4 digits.": Exit Function
                End If
            End If

        Case OP_EXP
            If Len(dcn) = 0 Then ValidateRow = "DCN is required.": Exit Function

        Case OP_TRF
            Dim st As String: st = UCase$(CellStr(ws, r, ColOf(cols, "status")))
            If Len(st) = 0 Then ValidateRow = "Status is required.": Exit Function
            If st <> "APPROVED" And st <> "REJECTED" Then
                ValidateRow = "Status must be APPROVED or REJECTED.": Exit Function
            End If
            If st = "REJECTED" Then
                If Len(CellStr(ws, r, ColOf(cols, "reason"))) = 0 Then
                    ValidateRow = "a Reason is required to reject.": Exit Function
                End If
            End If
    End Select
End Function

' Validate-pass detail: which optional fields carry a value, and which are
' being left out. Makes the omit-blanks rule visible before anything is sent.
'==============================================================================
' PRIVATE: what would actually change
'
' Validate used to check the SHAPE of a row - required fields present, values
' in range. It could not tell you that a row was about to overwrite a correct
' value with a stale one, because it never looked at what CCAT holds.
'
' That is the footgun this closes. Every Add/Update rewrites ownershipTypeCode,
' so a sheet pasted from last week's lookup silently reverts a record somebody
' fixed since. Nothing warned you. Now the Result cell says
' "OwnershipType RENTAL -> OWNED" and you get to read it before you Run.
'
' TWO RULES THIS FOLLOWS.
'
' It is ADVISORY ONLY. The diff never changes a row's verdict from OK to
' SKIPPED. Validate has always worked off the sheet alone, and making a row's
' fate depend on a live call means an API hiccup silently skips good rows. The
' note tells you an Expire has nothing to expire; deciding what to do about
' that stays yours.
'
' It NEVER FAILS THE VALIDATION. If the lookup errors, the row still validates
' and the note says the comparison was unavailable. Validate working while the
' proxy is down is a property worth keeping.
'==============================================================================

' Asks before spending a lookup per serial on a big sheet.
Private Function AskCompare(ByVal nRows As Long) As Boolean
    If nRows <= DIFF_PROMPT_ROWS Then AskCompare = True: Exit Function

    AskCompare = (MsgBox( _
        "Compare each row against what CCAT holds now?" & vbCrLf & vbCrLf & _
        "This is what shows you a row about to overwrite a good value with a " & _
        "stale one - but it costs one lookup per serial, and there are " & _
        nRows & " rows here." & vbCrLf & vbCrLf & _
        "Yes - slower, and says exactly what would change" & vbCrLf & _
        "No  - the usual field checks only", _
        vbQuestion + vbYesNo, "Cat Asset Tools - Validate") = vbYes)
End Function

Private Sub ResetDiffCache()
    Set mDiffCache = CreateObject("Scripting.Dictionary")
    mDiffCache.CompareMode = vbTextCompare
End Sub

' Every ownership record on a serial, as RecordValues arrays. Nothing when the
' lookup failed - cached either way, so a bad serial costs one call not ten.
Private Function RecordsFor(ByVal serial As String) As Collection
    If mDiffCache Is Nothing Then ResetDiffCache

    If mDiffCache.Exists(serial) Then
        If IsObject(mDiffCache(serial)) Then Set RecordsFor = mDiffCache(serial)
        Exit Function
    End If

    Dim txt As String, recs As Object, out As Collection, i As Long
    On Error GoTo Failed
    txt = CatSearch(serial, "")
    Set recs = OwnershipRecords(txt)
    On Error GoTo 0
    If recs Is Nothing Then GoTo Failed

    Set out = New Collection
    For i = 1 To recs.Count
        out.Add RecordValues(recs(i))
    Next i
    mDiffCache.Add serial, out
    Set RecordsFor = out
    Exit Function

Failed:
    ' Remember the miss so the next nine rows on this serial do not retry it.
    If Not mDiffCache.Exists(serial) Then mDiffCache.Add serial, Empty
End Function

' The record on this serial held by this DCN, if there is one.
Private Function FindByDcn(ByVal recs As Collection, ByVal dcn As String, _
                           ByRef found As Variant) As Boolean
    Dim i As Long, v As Variant
    For i = 1 To recs.Count
        v = recs(i)
        If StrComp(CStr(v(2)), dcn, vbTextCompare) = 0 Then
            found = v
            FindByDcn = True
            Exit Function
        End If
    Next i
End Function

' The state of this row's record immediately before it is sent, for the log.
'
' Returns FOUND / NONE / UNAVAILABLE and fills `before` on FOUND. Those three
' are kept apart deliberately: an Undo that cannot tell "there was no record"
' from "the lookup failed" would happily write blanks over a live one.
'
' Costs one lookup per distinct serial, shared with the Validate cache. That
' roughly doubles a Run's calls, and it buys the only route back from an
' Expire - which is a trade worth making without asking.
Private Function CaptureBefore(ByVal ws As Worksheet, ByVal r As Long, _
                               ByVal cols As Object, ByVal serial As String, _
                               ByRef before As Variant) As String
    before = Empty

    Dim recs As Collection
    Set recs = RecordsFor(serial)
    If recs Is Nothing Then CaptureBefore = BEFORE_UNAVAILABLE: Exit Function
    If recs.Count = 0 Then CaptureBefore = BEFORE_NONE: Exit Function

    Dim dcn As String: dcn = CleanId(CellStr(ws, r, ColOf(cols, "dcn")))

    ' Transfer sheets carry no DCN, and a serial can sit on several. With no
    ' DCN to match on, record the first - it is the whole record either way,
    ' and the log line keeps the serial so nothing is ambiguous later.
    Dim cur As Variant
    If Len(dcn) = 0 Then
        before = recs(1)
        CaptureBefore = BEFORE_FOUND
        Exit Function
    End If

    If FindByDcn(recs, dcn, cur) Then
        before = cur
        CaptureBefore = BEFORE_FOUND
    Else
        CaptureBefore = BEFORE_NONE
    End If
End Function

Private Function DiffNote(ByVal ws As Worksheet, ByVal r As Long, ByVal cols As Object, _
                          ByVal op As String, ByVal serial As String) As String
    Dim recs As Collection
    Set recs = RecordsFor(serial)
    If recs Is Nothing Then DiffNote = " - could not compare (lookup unavailable)": Exit Function

    If recs.Count = 0 Then
        Select Case op
            Case OP_ADD: DiffNote = " - NEW: this serial is not in CCAT at all"
            Case Else:   DiffNote = " - NOTHING TO DO: this serial is not in CCAT"
        End Select
        Exit Function
    End If

    Select Case op
        Case OP_ADD: DiffNote = DiffAddUpdate(ws, r, cols, recs)
        Case OP_EXP: DiffNote = DiffExpire(ws, r, cols, recs)
        Case OP_TRF: DiffNote = DiffTransfer(recs)
    End Select
End Function

' Only the fields this row will actually SEND are compared. A blank cell is
' omitted from the request entirely, so it cannot change anything and must not
' be reported as a change to blank.
'
' CustomAssetName is deliberately absent: the API returns assetName as a
' coalesced custom-or-base value, so there is no field to compare it against
' honestly, and a guess here would be worse than a gap.
Private Function DiffAddUpdate(ByVal ws As Worksheet, ByVal r As Long, _
                               ByVal cols As Object, ByVal recs As Collection) As String
    Dim dcn As String: dcn = CleanId(CellStr(ws, r, ColOf(cols, "dcn")))

    Dim cur As Variant
    If Not FindByDcn(recs, dcn, cur) Then
        DiffAddUpdate = " - NEW record for this DCN (serial exists on " & _
                        recs.Count & " other DCN(s))"
        Exit Function
    End If

    Dim keys As Variant, labels As Variant, idx As Variant
    keys = Array("ownershiptype", "model", "modelyear", "productfamilycode", _
                 "productfamilyname", "baseassetname")
    labels = Array("OwnershipType", "Model", "ModelYear", "ProductFamilyCode", _
                   "ProductFamilyName", "BaseAssetName")
    idx = Array(3, 4, 5, 18, 19, 21)

    Dim changes As String, i As Long, sheetVal As String, ccatVal As String
    For i = LBound(keys) To UBound(keys)
        sheetVal = CellStr(ws, r, ColOf(cols, CStr(keys(i))))
        If Len(sheetVal) > 0 Then
            ccatVal = CStr(cur(idx(i)))
            If StrComp(sheetVal, ccatVal, vbTextCompare) <> 0 Then
                changes = changes & IIf(Len(changes) > 0, ", ", "") & _
                          labels(i) & " " & IIf(Len(ccatVal) = 0, "(blank)", ccatVal) & _
                          " -> " & sheetVal
            End If
        End If
    Next i

    If Len(changes) = 0 Then
        DiffAddUpdate = " - no change"
    Else
        DiffAddUpdate = " - CHANGES: " & changes
    End If
End Function

' Expire removes an ownership record. Saying which one, in what state, is the
' last chance to notice it is not the one you meant.
Private Function DiffExpire(ByVal ws As Worksheet, ByVal r As Long, _
                            ByVal cols As Object, ByVal recs As Collection) As String
    Dim dcn As String: dcn = CleanId(CellStr(ws, r, ColOf(cols, "dcn")))

    Dim cur As Variant
    If Not FindByDcn(recs, dcn, cur) Then
        DiffExpire = " - NO record on DCN " & dcn & " - nothing to expire, this will fail"
        Exit Function
    End If

    DiffExpire = " - WILL EXPIRE " & CStr(cur(3)) & "/" & CStr(cur(6)) & _
                 " held by " & CStr(cur(11))
End Function

' Transfer acts on a pending request. If there isn't one, the send fails - and
' that is worth knowing before rather than after.
Private Function DiffTransfer(ByVal recs As Collection) As String
    Dim i As Long, v As Variant
    For i = 1 To recs.Count
        v = recs(i)
        If Len(Trim$(CStr(v(9)))) > 0 Then
            DiffTransfer = " - pending " & CStr(v(9)) & " on DCN " & CStr(v(2)) & _
                           " (" & CStr(v(14)) & ")"
            Exit Function
        End If
    Next i
    DiffTransfer = " - NO pending request on this serial, this will fail"
End Function

Private Function PreviewNote(ByVal ws As Worksheet, ByVal r As Long, ByVal cols As Object, _
                             ByVal op As String) As String
    If op <> OP_ADD Then Exit Function

    Dim names As Variant, keys As Variant
    names = Array("Model", "Model Year", "Product Family Code", "Product Family Name", _
                  "Base Asset Name", "Custom Asset Name")
    keys = Array("model", "modelyear", "productfamilycode", "productfamilyname", _
                 "baseassetname", "customassetname")

    Dim omitted As String, i As Long
    For i = LBound(keys) To UBound(keys)
        If Len(CellStr(ws, r, ColOf(cols, CStr(keys(i))))) = 0 Then
            omitted = omitted & IIf(Len(omitted) > 0, ", ", "") & names(i)
        End If
    Next i

    If Len(omitted) > 0 Then PreviewNote = " - omitting " & omitted
End Function

'==============================================================================
' PRIVATE: send one row. Returns the outcome text; kind 0=fail 1=ok.
'==============================================================================
Private Function SendRow(ByVal ws As Worksheet, ByVal r As Long, ByVal cols As Object, _
                         ByVal op As String, ByVal serial As String, _
                         ByRef kind As Long) As String
    Dim mk As String: mk = CellStr(ws, r, ColOf(cols, "makecode", "make"))
    Dim dmk As String: dmk = CellStr(ws, r, ColOf(cols, "dealermakecode", "dealermake"))
    Dim dcn As String: dcn = CleanId(CellStr(ws, r, ColOf(cols, "dcn")))

    Dim txt As String, status As Long, errMsg As String, gotResp As Boolean
    Dim attempt As Long
    gotResp = False

    For attempt = 1 To MAX_ATTEMPTS
        errMsg = "": status = 0
        On Error Resume Next
        Select Case op
            Case OP_ADD
                txt = CatAddUpdate(serial, dcn, mk, dmk, _
                        LCase$(CellStr(ws, r, ColOf(cols, "ownershiptype", "ownershiptypecode", "type"))), _
                        CellStr(ws, r, ColOf(cols, "model")), _
                        CellStr(ws, r, ColOf(cols, "modelyear", "year")), _
                        CellStr(ws, r, ColOf(cols, "productfamilycode", "pfcode")), _
                        CellStr(ws, r, ColOf(cols, "productfamilyname", "pfname")), _
                        CellStr(ws, r, ColOf(cols, "baseassetname", "basename")), _
                        CellStr(ws, r, ColOf(cols, "customassetname", "customname")), status)
            Case OP_EXP
                txt = CatExpire(serial, dcn, mk, dmk, status)
            Case OP_TRF
                txt = CatTransfer(serial, mk, dmk, _
                        UCase$(CellStr(ws, r, ColOf(cols, "status"))), _
                        CellStr(ws, r, ColOf(cols, "reason")), status)
        End Select
        If Err.Number <> 0 Then errMsg = Err.Description: Err.Clear
        On Error GoTo 0

        If Len(errMsg) = 0 Then gotResp = True: Exit For
        If Not IsTransient(errMsg) Then Exit For
        If attempt < MAX_ATTEMPTS Then Application.Wait Now + TimeSerial(0, 0, RETRY_WAIT_SECONDS)
    Next attempt

    If Not gotResp Then
        kind = 0
        SendRow = "FAILED: " & errMsg
        Exit Function
    End If

    If IsSuccess(op, status) Then
        kind = 1
        Dim st As String
        If op = OP_ADD Then st = FieldOf(txt, "status")
        SendRow = "OK (" & status & ")" & IIf(Len(st) > 0, " - " & st, "")
    Else
        kind = 0
        SendRow = "FAILED " & status & ": " & ProxyError(txt, status)
    End If
End Function

' Add/update answers 200 or 201; expire and transfer answer 204 (no content).
Private Function IsSuccess(ByVal op As String, ByVal status As Long) As Boolean
    Select Case op
        Case OP_ADD: IsSuccess = (status = 200 Or status = 201)
        Case Else:   IsSuccess = (status = 204 Or status = 200)
    End Select
End Function

'==============================================================================
' PRIVATE: sheet building
'==============================================================================
Private Sub BuildSheet(ByVal op As String)
    On Error GoTo Fail

    Dim wb As Workbook: Set wb = TargetBook()
    Dim nm As String, headers As Variant, notes As Variant, tiers As Variant
    SheetSpec op, nm, headers, notes, tiers

    Dim ws As Worksheet, shp As Shape
    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    On Error GoTo Fail
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = nm
    Else
        If MsgBox("'" & nm & "' already exists. Clear it and rebuild?" & vbCrLf & vbCrLf & _
                  "Anything on it will be lost.", vbExclamation + vbYesNo + vbDefaultButton2, _
                  "Cat Asset Tools") <> vbYes Then Exit Sub
        ws.Cells.Clear
        For Each shp In ws.Shapes: shp.Delete: Next shp
    End If

    Dim c As Long, bg As Long, fg As Long
    For c = 0 To UBound(headers)
        Select Case tiers(c)
            Case 0: bg = RGB(31, 78, 121): fg = vbWhite            ' always required
            Case 1: bg = RGB(46, 117, 182): fg = vbWhite           ' conditional
            Case 3: bg = RGB(89, 89, 89): fg = vbWhite             ' output
            Case Else: bg = RGB(189, 215, 238): fg = RGB(31, 31, 31)  ' optional
        End Select
        With ws.Cells(HEADER_ROW, c + 1)
            .Value = headers(c)
            .Font.Bold = True
            .Font.Color = fg
            .Interior.Color = bg
            If Not .Comment Is Nothing Then .Comment.Delete
            .AddComment CStr(notes(c))
            .Comment.Shape.TextFrame.AutoSize = True
            .Comment.Visible = False
        End With
    Next c

    ' Dropdowns on the data rows
    If op = OP_ADD Then
        AddList ws, ColOfHeader(headers, "Ownership Type"), _
                "owned,rental,leased,sold,inventory,unknown"
    ElseIf op = OP_TRF Then
        AddList ws, ColOfHeader(headers, "Status"), "APPROVED,REJECTED"
    End If

    ws.Range("A1").AutoFilter
    ws.Activate
    ws.Rows(HEADER_ROW + 1).Select
    ActiveWindow.FreezePanes = True
    ws.Columns.AutoFit

    With ws.Cells(3, UBound(headers) + 3)
        .Value = OpLabel(op) & vbLf & vbLf & _
                 "Fill one row per asset - a single asset is just one row." & vbLf & _
                 "Then click Validate, read the Result column, then Run." & vbLf & vbLf & _
                 "The first columns are in the same order as the batch-lookup" & vbLf & _
                 "output, so you can Ctrl-select those columns on a results" & vbLf & _
                 "sheet, copy, and paste them straight in here." & vbLf & vbLf & _
                 "- Dark blue  = always required" & vbLf & _
                 "- Mid blue   = conditional (see the header comment)" & vbLf & _
                 "- Light blue = optional; blank means CCAT keeps its value" & vbLf & _
                 "- Grey       = written by the macro; do not edit" & vbLf & vbLf & _
                 "Hover any header for details."
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    ws.Range(ws.Cells(3, UBound(headers) + 3), ws.Cells(14, UBound(headers) + 7)).Merge

    ws.Cells(HEADER_ROW + 1, 1).Select
    MsgBox "'" & nm & "' is ready." & vbCrLf & vbCrLf & _
           "Fill rows, then CCAT > Validate, then Run.", _
           vbInformation, "Cat Asset Tools"
    Exit Sub
Fail:
    MsgBox "Error: " & Err.Description, vbCritical, "Cat Asset Tools"
End Sub

' Column layouts. The leading block of each mirrors the batch-lookup column
' order (see the module header) so a multi-column copy pastes straight in.
Private Sub SheetSpec(ByVal op As String, ByRef nm As String, ByRef headers As Variant, _
                      ByRef notes As Variant, ByRef tiers As Variant)
    Select Case op
        Case OP_ADD
            nm = SH_ADD
            ' Columns 1-6 are the paste block and MUST stay in this order and
            ' adjacent - they mirror HeaderArray columns 1-6. Dealer Make Code
            ' is the rarely-used alternative to Make Code, so it sits at the far
            ' end rather than splitting the block.
            headers = Array("Serial", "Make Code", "DCN", "Ownership Type", _
                            "Model", "Model Year", _
                            "Product Family Code", "Product Family Name", _
                            "Base Asset Name", "Custom Asset Name", _
                            "Dealer Make Code", "Result")
            tiers = Array(0, 0, 0, 0, 1, 1, 2, 2, 2, 2, 2, 3)
            notes = Array( _
                "REQUIRED. Asset serial number (exact match).", _
                "REQUIRED. Caterpillar manufacturer code, e.g. CAT or CW1. This is the one to use - it comes back on every lookup and means the same thing to everyone.", _
                "REQUIRED. Dealer Customer Number.", _
                "REQUIRED on every row. One of: owned, rental, leased, sold, inventory, unknown. The API only demands it for a NEW record, but a blank here is the most common cause of a rejection, so it is enforced.", _
                "Required for a NEW record. Asset model, e.g. 980H. Max 65 characters. Blank on an existing record leaves the stored model alone.", _
                "Required for a NEW record. 4-digit year of manufacture, e.g. 2006. Blank on an existing record leaves the stored year alone.", _
                "Optional. Cat product family code, e.g. MDWL. Max 50.", _
                "Optional. Product family name, e.g. MEDIUM WHEEL LOADER. Max 50.", _
                "Optional. Canonical asset name set by the dealer. Max 60.", _
                "Optional. Your own label; shown in preference to Base Asset Name. Max 60.", _
                "OPTIONAL - the alternative to Make Code, not an extra. Your dealership's own two-character code, e.g. AA. Leave it blank and use Make Code; fill this ONLY when your source data has the dealer code and not the Cat one. Supplying both is rejected.", _
                "Written by the macro: OK / FAILED / SKIPPED. Do not edit.")

        Case OP_EXP
            nm = SH_EXP
            headers = Array("Serial", "Make Code", "DCN", "Dealer Make Code", "Result")
            tiers = Array(0, 0, 0, 2, 3)
            notes = Array( _
                "REQUIRED. Asset serial number (exact match).", _
                "REQUIRED. Caterpillar manufacturer code, e.g. CAT or CW1. This is the one to use.", _
                "REQUIRED. Dealer Customer Number of the record to remove.", _
                "OPTIONAL - the alternative to Make Code, not an extra. e.g. AA. Leave blank and use Make Code. Supplying both is rejected.", _
                "Written by the macro: OK / FAILED / SKIPPED. Do not edit.")

        Case OP_TRF
            nm = SH_TRF
            ' Columns 1-2 are the paste block (HeaderArray 1-2). Status and
            ' Reason are decisions you type, not values you look up, so they
            ' follow; Dealer Make Code goes last for the same reason as above.
            headers = Array("Serial", "Make Code", "Status", "Reason", _
                            "Dealer Make Code", "Result")
            tiers = Array(0, 0, 0, 1, 2, 3)
            notes = Array( _
                "REQUIRED. Asset serial number (exact match).", _
                "REQUIRED. Caterpillar manufacturer code, e.g. CAT or CW1. This is the one to use.", _
                "REQUIRED. APPROVED or REJECTED. No DCN is used by this endpoint. Filter a lookup on OwnershipRequestType = RECEIVED to find the ones waiting on you.", _
                "Required only when Status is REJECTED; optional otherwise.", _
                "OPTIONAL - the alternative to Make Code, not an extra. e.g. AA. Leave blank and use Make Code. Supplying both is rejected.", _
                "Written by the macro: OK / FAILED / SKIPPED. Do not edit.")
    End Select
End Sub

Private Function ColOfHeader(ByVal headers As Variant, ByVal label As String) As Long
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        If headers(i) = label Then ColOfHeader = i + 1: Exit Function
    Next i
End Function

Private Sub AddList(ByVal ws As Worksheet, ByVal col As Long, ByVal listCsv As String)
    If col = 0 Then Exit Sub
    With ws.Range(ws.Cells(HEADER_ROW + 1, col), ws.Cells(HEADER_ROW + 1000, col)).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:=listCsv
    End With
End Sub

'==============================================================================
' PRIVATE: small helpers (carried over from CatBatchActions unchanged)
'==============================================================================
Private Function EnsureResultColumn(ByVal ws As Worksheet, ByVal cols As Object) As Long
    EnsureResultColumn = ColOf(cols, "result")
    If EnsureResultColumn = 0 Then
        EnsureResultColumn = LastHeaderCol(ws) + 1
        ws.Cells(HEADER_ROW, EnsureResultColumn).Value = "Result"
        ws.Cells(HEADER_ROW, EnsureResultColumn).Font.Bold = True
    End If
End Function

Private Function CountRows(ByVal ws As Worksheet, ByVal cSerial As Long, _
                           ByVal lastRow As Long) As Long
    Dim r As Long
    For r = HEADER_ROW + 1 To lastRow
        If Len(CleanId(CStr(ws.Cells(r, cSerial).Value))) > 0 Then CountRows = CountRows + 1
    Next r
End Function

Private Function HeaderMap(ByVal ws As Worksheet) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim lastCol As Long: lastCol = LastHeaderCol(ws)
    Dim c As Long, h As String
    For c = 1 To lastCol
        h = NormHeader(CStr(ws.Cells(HEADER_ROW, c).Value))
        If Len(h) > 0 And Not d.Exists(h) Then d(h) = c
    Next c
    Set HeaderMap = d
End Function

Private Function LastHeaderCol(ByVal ws As Worksheet) As Long
    LastHeaderCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
End Function

Private Function ColOf(ByVal d As Object, ParamArray candidates() As Variant) As Long
    Dim i As Long
    For i = LBound(candidates) To UBound(candidates)
        If d.Exists(CStr(candidates(i))) Then ColOf = d(CStr(candidates(i))): Exit Function
    Next i
    ColOf = 0
End Function

Private Function NormHeader(ByVal s As String) As String
    s = LCase$(Trim$(s))
    s = Replace(s, " ", ""): s = Replace(s, "_", ""): s = Replace(s, "-", "")
    s = Replace(s, "/", ""): s = Replace(s, ".", "")
    NormHeader = s
End Function

Private Function CellStr(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As String
    If c = 0 Then Exit Function
    CellStr = Trim$(CStr(ws.Cells(r, c).Value))
End Function

' kind: 0 = fail (red), 1 = ok (green), 2 = skipped (yellow)
Private Sub WriteResult(ByVal ws As Worksheet, ByVal r As Long, ByVal cResult As Long, _
                        ByVal msg As String, ByVal kind As Long)
    With ws.Cells(r, cResult)
        .Value = msg
        Select Case kind
            Case 1: .Interior.Color = RGB(226, 242, 226)
            Case 2: .Interior.Color = RGB(255, 242, 204)
            Case Else: .Interior.Color = RGB(255, 220, 220)
        End Select
    End With
End Sub

Private Function IsTransient(ByVal msg As String) As Boolean
    Dim m As String: m = LCase$(msg)
    If InStr(m, "429") > 0 Then IsTransient = True: Exit Function
    If InStr(m, "api 5") > 0 Then IsTransient = True: Exit Function
    If InStr(m, "timed out") > 0 Or InStr(m, "timeout") > 0 Then IsTransient = True: Exit Function
    If InStr(m, "could not be resolved") > 0 Or InStr(m, "cannot connect") > 0 _
       Or InStr(m, "connection") > 0 Or InStr(m, "winhttp") > 0 Then IsTransient = True: Exit Function
    IsTransient = False
End Function
