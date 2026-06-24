Attribute VB_Name = "CatBatchActions"
'==============================================================================
' Cat Asset Management V2 - BATCH add / update macro (via PROXY)
'
' The write counterpart to CatBatchLookup: reads a sheet of equipment rows and
' sends each one to POST /api/ownership through the proxy (add-or-update), then
' writes a per-row outcome back to a "Result" column. No credentials live in the
' workbook - only the proxy URL + function key (Config sheet).
'
' Run CatSetupBatchAddUpdateSheet once to build the "Batch Add-Update" input
' sheet with the right column headers, a dropdown, and a Run button. Fill one
' row per asset, then click Run (or run CatBatchAddUpdate).
'
' Columns (header row 1; matched by name, order/extra columns don't matter):
'   Serial*        MakeCode / DealerMakeCode (one)*   DCN*
'   OwnershipType  Model  ModelYear  ProductFamilyCode  ProductFamilyName
'   BaseAssetName  CustomAssetName    Result (written by the macro)
'   (* required on every row. A NEW record also needs OwnershipType, Model,
'      ModelYear - the API rejects it otherwise.)
'
' Depends on the CatAssetLookup module (CatAddUpdate, CleanId, FieldOf,
' ProxyError) and VBA-JSON.
'
' WARNING: this changes real data when the proxy has CAT_ENABLE_WRITES=true.
' While writes are disabled (the default) every row comes back "403
' writes_disabled" and nothing changes. Keep the proxy pinned to a TEST dealer
' code (CAT_FORCE_PARTY_NUMBER) until you intend to touch production.
'==============================================================================
Option Explicit

Private Const INPUT_SHEET As String = "Batch Add-Update"
Private Const HEADER_ROW As Long = 1
Private Const MAX_ATTEMPTS As Long = 3        ' retry transient failures only
Private Const RETRY_WAIT_SECONDS As Long = 1

'==============================================================================
' Run: process every data row on the active (or "Batch Add-Update") sheet
'==============================================================================
Public Sub CatBatchAddUpdate()
    On Error GoTo Fail

    Dim ws As Worksheet: Set ws = TargetSheet()
    If ws Is Nothing Then Exit Sub

    ' Map header names -> column numbers
    Dim cols As Object: Set cols = HeaderMap(ws)
    Dim cSerial As Long: cSerial = ColOf(cols, "serial", "serialnumber")
    If cSerial = 0 Then
        MsgBox "No 'Serial' column found in row " & HEADER_ROW & " of '" & ws.Name & "'." & vbCrLf & _
               "Run CatSetupBatchAddUpdateSheet to build the input sheet.", vbExclamation
        Exit Sub
    End If

    Dim cMake As Long: cMake = ColOf(cols, "makecode", "make")
    Dim cDmk As Long: cDmk = ColOf(cols, "dealermakecode", "dealermake")
    Dim cDcn As Long: cDcn = ColOf(cols, "dcn")
    Dim cType As Long: cType = ColOf(cols, "ownershiptype", "ownershiptypecode", "type")
    Dim cModel As Long: cModel = ColOf(cols, "model")
    Dim cYear As Long: cYear = ColOf(cols, "modelyear", "year")
    Dim cPfc As Long: cPfc = ColOf(cols, "productfamilycode", "pfcode")
    Dim cPfn As Long: cPfn = ColOf(cols, "productfamilyname", "pfname")
    Dim cBase As Long: cBase = ColOf(cols, "baseassetname", "basename")
    Dim cCustom As Long: cCustom = ColOf(cols, "customassetname", "customname")

    ' Result column - create one if the sheet doesn't have it yet
    Dim cResult As Long: cResult = ColOf(cols, "result")
    If cResult = 0 Then
        cResult = LastHeaderCol(ws) + 1
        ws.Cells(HEADER_ROW, cResult).Value = "Result"
        ws.Cells(HEADER_ROW, cResult).Font.Bold = True
    End If

    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, cSerial).End(xlUp).Row
    If lastRow <= HEADER_ROW Then MsgBox "No data rows under the headers.", vbExclamation: Exit Sub

    ' Count the rows that actually have a serial
    Dim r As Long, nRows As Long
    For r = HEADER_ROW + 1 To lastRow
        If Len(CleanId(CStr(ws.Cells(r, cSerial).Value))) > 0 Then nRows = nRows + 1
    Next r
    If nRows = 0 Then MsgBox "No serial numbers found under the headers.", vbExclamation: Exit Sub

    If MsgBox(nRows & " row(s) will be sent to the proxy as ADD / UPDATE requests." & vbCrLf & vbCrLf & _
              "This changes real data if the proxy has writes enabled." & vbCrLf & _
              "Continue?", vbExclamation + vbYesNo + vbDefaultButton2, "Confirm Batch Add / Update") <> vbYes Then Exit Sub

    Dim nOk As Long, nFailed As Long, nSkipped As Long
    Application.ScreenUpdating = False

    Dim done As Long
    For r = HEADER_ROW + 1 To lastRow
        Dim serial As String: serial = CleanId(CStr(ws.Cells(r, cSerial).Value))
        If Len(serial) = 0 Then GoTo NextRow                ' blank row inside the range

        done = done + 1
        Application.StatusBar = "Add/Update " & done & " of " & nRows & "  (" & serial & ")"

        Dim dcn As String: dcn = CleanId(CellStr(ws, r, cDcn))
        Dim mk As String: mk = CellStr(ws, r, cMake)
        Dim dmk As String: dmk = CellStr(ws, r, cDmk)

        ' --- per-row validation (mirrors the single-asset button) ---
        If Len(serial) = 0 Or Len(dcn) = 0 Then
            WriteResult ws, r, cResult, "SKIPPED: Serial and DCN are required.", 2
            nSkipped = nSkipped + 1: GoTo NextRow
        End If
        If Len(mk) = 0 And Len(dmk) = 0 Then
            WriteResult ws, r, cResult, "SKIPPED: provide Make Code or Dealer Make Code.", 2
            nSkipped = nSkipped + 1: GoTo NextRow
        End If

        ' --- call with retry for transient failures ---
        Dim txt As String, status As Long, errMsg As String, gotResp As Boolean
        Dim attempt As Long
        gotResp = False: errMsg = ""
        For attempt = 1 To MAX_ATTEMPTS
            errMsg = "": status = 0
            On Error Resume Next
            txt = CatAddUpdate(serial, dcn, mk, dmk, LCase$(CellStr(ws, r, cType)), _
                               CellStr(ws, r, cModel), CellStr(ws, r, cYear), _
                               CellStr(ws, r, cPfc), CellStr(ws, r, cPfn), _
                               CellStr(ws, r, cBase), CellStr(ws, r, cCustom), status)
            If Err.Number <> 0 Then errMsg = Err.Description: Err.Clear
            On Error GoTo Fail
            If Len(errMsg) = 0 Then gotResp = True: Exit For
            If Not IsTransient(errMsg) Then Exit For
            If attempt < MAX_ATTEMPTS Then Application.Wait Now + TimeSerial(0, 0, RETRY_WAIT_SECONDS)
        Next attempt

        ' --- classify + record the outcome ---
        If Not gotResp Then
            WriteResult ws, r, cResult, "FAILED: " & errMsg, 0
            nFailed = nFailed + 1
        ElseIf status = 200 Or status = 201 Then
            Dim st As String: st = FieldOf(txt, "status")
            WriteResult ws, r, cResult, "OK (" & status & ")" & IIf(Len(st) > 0, " - " & st, ""), 1
            nOk = nOk + 1
        Else
            WriteResult ws, r, cResult, "FAILED " & status & ": " & ProxyError(txt, status), 0
            nFailed = nFailed + 1
        End If
NextRow:
    Next r

    ws.Columns.AutoFit
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox done & " row(s) processed:" & vbCrLf & _
           "  - " & nOk & " OK (added / updated)" & vbCrLf & _
           "  - " & nFailed & " failed (red - safe to re-run; add/update is idempotent)" & vbCrLf & _
           "  - " & nSkipped & " skipped (yellow - missing required fields)", _
           IIf(nFailed > 0, vbExclamation, vbInformation), "Batch Add / Update"
    Exit Sub
Fail:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical, "Batch Add / Update"
End Sub

'==============================================================================
' Helpers
'==============================================================================

' Use the dedicated input sheet if it exists, otherwise the active sheet.
Private Function TargetSheet() As Worksheet
    On Error Resume Next
    Set TargetSheet = ThisWorkbook.Worksheets(INPUT_SHEET)
    On Error GoTo 0
    If TargetSheet Is Nothing Then Set TargetSheet = ActiveSheet
End Function

' Build a dictionary: normalized header text -> column number (header row).
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

' Return the column for the first matching candidate header (0 if none).
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

' Trimmed string from a cell, or "" when the column isn't present (col = 0).
Private Function CellStr(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As String
    If c = 0 Then Exit Function
    CellStr = Trim$(CStr(ws.Cells(r, c).Value))
End Function

' Write the outcome + shade the row's result cell. kind: 0=fail(red) 1=ok(green) 2=skip(yellow)
Private Sub WriteResult(ByVal ws As Worksheet, ByVal r As Long, ByVal cResult As Long, _
                        ByVal msg As String, ByVal kind As Long)
    With ws.Cells(r, cResult)
        .Value = msg
        Select Case kind
            Case 1: .Interior.Color = RGB(226, 242, 226)    ' green
            Case 2: .Interior.Color = RGB(255, 242, 204)    ' yellow
            Case Else: .Interior.Color = RGB(255, 220, 220) ' red
        End Select
    End With
End Sub

' Same transient/permanent split as the batch lookup.
Private Function IsTransient(ByVal msg As String) As Boolean
    Dim m As String: m = LCase$(msg)
    If InStr(m, "429") > 0 Then IsTransient = True: Exit Function
    If InStr(m, "api 5") > 0 Then IsTransient = True: Exit Function
    If InStr(m, "timed out") > 0 Or InStr(m, "timeout") > 0 Then IsTransient = True: Exit Function
    If InStr(m, "could not be resolved") > 0 Or InStr(m, "cannot connect") > 0 _
       Or InStr(m, "connection") > 0 Or InStr(m, "winhttp") > 0 Then IsTransient = True: Exit Function
    IsTransient = False
End Function

'==============================================================================
' One-time setup: build the "Batch Add-Update" input sheet
'==============================================================================
Public Sub CatSetupBatchAddUpdateSheet()
    Dim ws As Worksheet, shp As Shape
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(INPUT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = INPUT_SHEET
    Else
        ws.Cells.Clear
        For Each shp In ws.Shapes: shp.Delete: Next shp
    End If

    ' Friendly header text (spaces are stripped when matching, so these still map
    ' to the API fields), plus a hover comment per column sourced from the spec.
    Dim headers As Variant, notes As Variant
    headers = Array("Serial", "Make Code", "Dealer Make Code", "DCN", "Ownership Type", _
                    "Model", "Model Year", "Product Family Code", "Product Family Name", _
                    "Base Asset Name", "Custom Asset Name", "Result")
    notes = Array( _
        "REQUIRED. Asset serial number (exact match). Identify the asset together with Make Code or Dealer Make Code.", _
        "REQUIRED (this OR Dealer Make Code, not both). Caterpillar manufacturer code, e.g. CW1.", _
        "REQUIRED (this OR Make Code, not both). Dealer-specific make code, usually 2 chars, e.g. CW.", _
        "REQUIRED. Dealer Customer Number - the dealer/customer association for the record.", _
        "Required for a NEW record. One of: owned, rental, leased, sold, inventory, unknown.", _
        "Required for a NEW record. Asset model, e.g. 980H. Max 65 characters.", _
        "Required for a NEW record. 4-digit year; for heavy equipment this is the year of manufacture, e.g. 2006.", _
        "Optional. Cat product family code, e.g. MDWL. Max 50 characters.", _
        "Optional. Product family name, e.g. MEDIUM WHEEL LOADER. Max 50 characters.", _
        "Optional. Canonical asset name set by the dealer, e.g. ABC123. Max 60 characters.", _
        "Optional. Custom asset name, e.g. Excavator #1. Shown in preference to Base Asset Name. Max 60 characters.", _
        "Filled by the macro per row: OK / FAILED / SKIPPED. Do not edit.")

    Dim c As Long, bg As Long, fg As Long
    For c = 0 To UBound(headers)
        Select Case c
            Case 0 To 3: bg = RGB(31, 78, 121): fg = vbWhite        ' always required
            Case 4 To 6: bg = RGB(46, 117, 182): fg = vbWhite       ' required for a NEW record
            Case 11: bg = RGB(89, 89, 89): fg = vbWhite             ' output (do not edit)
            Case Else: bg = RGB(189, 215, 238): fg = RGB(31, 31, 31) ' optional
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

    ' Ownership Type dropdown on the data rows
    Dim typeCol As Long: typeCol = 5      ' "OwnershipType"
    With ws.Range(ws.Cells(HEADER_ROW + 1, typeCol), ws.Cells(HEADER_ROW + 1000, typeCol)).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="owned,rental,leased,sold,inventory,unknown"
    End With

    ws.Range("A1").AutoFilter
    ws.Activate
    ws.Rows(HEADER_ROW + 1).Select
    ActiveWindow.FreezePanes = True
    ws.Columns.AutoFit

    ' Run button
    Dim t As Double: t = ws.Cells(1, UBound(headers) + 3).Top
    Dim l As Double: l = ws.Cells(1, UBound(headers) + 3).Left
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, l, t, 170, 28)
    btn.OnAction = "CatBatchAddUpdate"
    btn.Fill.ForeColor.RGB = RGB(31, 78, 121)
    btn.Line.Visible = msoFalse
    With btn.TextFrame2.TextRange
        .Text = "Run Batch Add / Update"
        .Font.Fill.ForeColor.RGB = vbWhite
        .Font.Bold = msoTrue
    End With

    ' Guidance note (legend matches the header colour tiers)
    With ws.Cells(3, UBound(headers) + 3)
        .Value = "Fill one row per asset, then click Run." & vbLf & _
                 "- Dark blue = always required: Serial, Make Code OR Dealer Make Code, DCN." & vbLf & _
                 "- Medium blue = required for a NEW asset: Ownership Type, Model, Model Year." & vbLf & _
                 "- Light blue = optional." & vbLf & _
                 "- Grey (Result) = filled by the macro; do not edit." & vbLf & _
                 "Hover any header for details. Writes only take effect when the proxy has CAT_ENABLE_WRITES=true."
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    ws.Range(ws.Cells(3, UBound(headers) + 3), ws.Cells(8, UBound(headers) + 6)).Merge

    ws.Cells(HEADER_ROW + 1, 1).Select
    MsgBox "'" & INPUT_SHEET & "' sheet is ready. Fill rows, then click Run Batch Add / Update.", _
           vbInformation, "Cat Asset Management"
End Sub
