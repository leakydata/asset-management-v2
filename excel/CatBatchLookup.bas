Attribute VB_Name = "CatBatchLookup"
'==============================================================================
' Cat Asset Management V2 - batch lookup macro
'
' Prompts you to select a column of serial numbers, looks each one up via the
' Cat Asset Management V2 API, and writes a fresh results sheet with one row per
' ownership record. Calls into the CatAssetLookup module (CatSearch, HeaderArray,
' RecordValues) so both modules stay consistent and share one cached token.
'
' Each input serial is classified into exactly one outcome:
'   * found        -> one row per ownership record (Note blank)
'   * NOT IN CCAT  -> the API returned successfully with zero records
'   * LOOKUP FAILED-> the call errored / timed out / gave an unexpected response.
'                     These are NEVER recorded as "not in CCAT", because we don't
'                     actually know whether the asset exists. Re-run failed ones.
'
' Run from Developer > Macros, or assign CatBatchLookup to a button/shape.
'==============================================================================
Option Explicit

Private Const MAX_ATTEMPTS As Long = 3        ' retry transient failures
Private Const RETRY_WAIT_SECONDS As Long = 1

Public Sub CatBatchLookup()
    On Error GoTo Fail

    ' 1) Ask the user to select the column of serial numbers
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox( _
        Prompt:="Select the column of serial numbers (a header row is OK):", _
        Title:="Cat Asset Lookup", Type:=8)
    On Error GoTo Fail
    If rng Is Nothing Then Exit Sub                 ' user cancelled

    ' Limit to the used area so selecting a whole column isn't a million cells
    Set rng = Application.Intersect(rng.Columns(1), rng.Worksheet.UsedRange)
    If rng Is Nothing Then MsgBox "That selection has no data.", vbExclamation: Exit Sub

    ' 2) Collect serials (skip blanks and an obvious header cell)
    Dim serials As Collection: Set serials = New Collection
    Dim cell As Range, s As String
    For Each cell In rng.Cells
        s = Trim$(CStr(cell.Value))
        If Len(s) > 0 Then
            If Not (serials.Count = 0 And LCase$(s) Like "*serial*") Then serials.Add s
        End If
    Next cell
    If serials.Count = 0 Then MsgBox "No serial numbers found in the selection.", vbExclamation: Exit Sub

    ' 3) Loop the serials, accumulate result rows
    Dim headers As Variant: headers = HeaderArray()
    Dim nFields As Long: nFields = UBound(headers) + 1
    Dim totalCols As Long: totalCols = 1 + nFields + 1     ' QuerySerial + fields + Note

    Dim rowsOut As Collection: Set rowsOut = New Collection
    Dim nFound As Long, nNotInCcat As Long, nFailed As Long
    Application.ScreenUpdating = False

    Dim idx As Long
    For idx = 1 To serials.Count
        s = serials(idx)
        Application.StatusBar = "Looking up " & idx & " of " & serials.Count & "  (" & s & ")"

        ' --- call with retry for transient failures ---
        Dim respText As String, errMsg As String, gotResp As Boolean
        Dim attempt As Long
        gotResp = False: errMsg = ""
        For attempt = 1 To MAX_ATTEMPTS
            errMsg = ""
            On Error Resume Next
            respText = CatSearch(s, "")
            If Err.Number <> 0 Then errMsg = Err.Description: Err.Clear
            On Error GoTo Fail
            If Len(errMsg) = 0 Then gotResp = True: Exit For
            If Not IsTransient(errMsg) Then Exit For           ' permanent error -> stop
            If attempt < MAX_ATTEMPTS Then Application.Wait Now + TimeSerial(0, 0, RETRY_WAIT_SECONDS)
        Next attempt

        ' --- classify the outcome ---
        If Not gotResp Then
            ' genuine failure - we do NOT know if the asset exists
            rowsOut.Add MakeRow(s, Nothing, "LOOKUP FAILED: " & errMsg, totalCols, nFields)
            nFailed = nFailed + 1
        Else
            Dim root As Object, recs As Object
            Set root = Nothing: Set recs = Nothing
            On Error Resume Next
            Set root = JsonConverter.ParseJson(respText)
            On Error GoTo Fail

            If root Is Nothing Then
                rowsOut.Add MakeRow(s, Nothing, "LOOKUP FAILED: unreadable response", totalCols, nFields)
                nFailed = nFailed + 1
            ElseIf Not root.Exists("ownershipRecords") Then
                rowsOut.Add MakeRow(s, Nothing, "LOOKUP FAILED: unexpected response", totalCols, nFields)
                nFailed = nFailed + 1
            Else
                Set recs = root("ownershipRecords")
                If recs.Count = 0 Then
                    ' confirmed: API answered, asset not present
                    rowsOut.Add MakeRow(s, Nothing, "NOT IN CCAT", totalCols, nFields)
                    nNotInCcat = nNotInCcat + 1
                Else
                    Dim k As Long
                    For k = 1 To recs.Count
                        rowsOut.Add MakeRow(s, recs(k), "", totalCols, nFields)
                    Next k
                    nFound = nFound + 1
                End If
            End If
        End If
    Next idx

    ' 4) Write everything to a fresh sheet in one shot
    Dim ws As Worksheet: Set ws = MakeOutputSheet()
    Dim grid() As Variant
    ReDim grid(1 To rowsOut.Count + 1, 1 To totalCols)

    grid(1, 1) = "QuerySerial"
    Dim c As Long
    For c = 0 To nFields - 1: grid(1, 2 + c) = headers(c): Next c
    grid(1, totalCols) = "Note"

    Dim rr As Long, rowArr As Variant
    For rr = 1 To rowsOut.Count
        rowArr = rowsOut(rr)
        For c = 0 To totalCols - 1: grid(rr + 1, c + 1) = rowArr(c): Next c
    Next rr

    ws.Range(ws.Cells(1, 1), ws.Cells(rowsOut.Count + 1, totalCols)).Value = grid
    ws.Rows(1).Font.Bold = True
    ws.Range("A1").AutoFilter
    HighlightOutcomes ws, rowsOut.Count, totalCols
    ws.Columns.AutoFit
    ws.Activate

    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox serials.Count & " serial(s) processed:" & vbCrLf & _
           "  - " & nFound & " found (" & rowsOut.Count - nNotInCcat - nFailed & " record rows)" & vbCrLf & _
           "  - " & nNotInCcat & " not in CCAT" & vbCrLf & _
           "  - " & nFailed & " lookup failed (highlighted - safe to re-run)", _
           IIf(nFailed > 0, vbExclamation, vbInformation), "Cat Asset Lookup"
    Exit Sub
Fail:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical, "Cat Asset Lookup"
End Sub

'==============================================================================
' Helpers (private to this module)
'==============================================================================

' Decide whether an error is worth retrying. 4xx like 403 are permanent;
' 429, 5xx, timeouts and connection errors are transient.
Private Function IsTransient(ByVal msg As String) As Boolean
    Dim m As String: m = LCase$(msg)
    If InStr(m, "429") > 0 Then IsTransient = True: Exit Function
    If InStr(m, "api 5") > 0 Or InStr(m, "token 5") > 0 Then IsTransient = True: Exit Function
    If InStr(m, "timed out") > 0 Or InStr(m, "timeout") > 0 Then IsTransient = True: Exit Function
    If InStr(m, "could not be resolved") > 0 Or InStr(m, "cannot connect") > 0 _
       Or InStr(m, "connection") > 0 Or InStr(m, "winhttp") > 0 Then IsTransient = True: Exit Function
    IsTransient = False
End Function

' Builds one output row: QuerySerial + 14 fields (if rec given) + Note.
Private Function MakeRow(ByVal serial As String, ByVal rec As Object, _
                         ByVal note As String, ByVal totalCols As Long, _
                         ByVal nFields As Long) As Variant
    Dim r() As Variant
    ReDim r(0 To totalCols - 1)
    r(0) = serial
    If Not rec Is Nothing Then
        Dim vals As Variant: vals = RecordValues(rec)
        Dim c As Long
        For c = 0 To nFields - 1: r(1 + c) = vals(c): Next c
    End If
    r(totalCols - 1) = note
    MakeRow = r
End Function

' Light row shading: red for failures, grey for not-in-CCAT.
Private Sub HighlightOutcomes(ByVal ws As Worksheet, ByVal nRows As Long, ByVal totalCols As Long)
    Dim r As Long, note As String
    For r = 1 To nRows
        note = CStr(ws.Cells(r + 1, totalCols).Value)
        If Left$(note, 13) = "LOOKUP FAILED" Then
            ws.Range(ws.Cells(r + 1, 1), ws.Cells(r + 1, totalCols)).Interior.Color = RGB(255, 220, 220)
        ElseIf note = "NOT IN CCAT" Then
            ws.Range(ws.Cells(r + 1, 1), ws.Cells(r + 1, totalCols)).Interior.Color = RGB(242, 242, 242)
        End If
    Next r
End Sub

'==============================================================================
' Output sheet
'==============================================================================
Private Function MakeOutputSheet() As Worksheet
    Dim base As String: base = "Asset Lookup Results"
    Dim nm As String: nm = base
    Dim i As Long: i = 1
    Do While SheetExists(nm)
        i = i + 1
        nm = base & " " & i
    Loop
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = nm
    Set MakeOutputSheet = ws
End Function

Private Function SheetExists(ByVal nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function
