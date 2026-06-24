Attribute VB_Name = "CatActions"
'==============================================================================
' Cat Asset Management V2 - actions (add/update, expire, transfer, check) via PROXY
'
' These call WRITE endpoints on the Asset Management PROXY, so they are deliberate
' macros (never worksheet functions) and each asks for confirmation. Run
' CatSetupActionsSheet once to build an "Actions" sheet with input cells and
' buttons. No credentials live in the workbook - only the proxy URL + function
' key (Config sheet).
'
' The PROXY decides whether writes are allowed: if its CAT_ENABLE_WRITES setting
' is false (the default), the write buttons return "403 writes_disabled". The
' Check button (read) always works.
'
' Depends on the CatAssetLookup module (CatAddUpdate, CatExpire, CatTransfer,
' CatSearch, OwnershipRecords, RecordValues, FieldOf, ProxyError) and VBA-JSON.
'
' WARNING: add / expire / transfer change real data. The proxy acts on the dealer
' code it is configured with - keep that on a test dealer code and use test assets
' until you intend to touch production.
'==============================================================================
Option Explicit

Private Const ACTIONS_SHEET As String = "Actions"

'==============================================================================
' Action buttons
'==============================================================================
Public Sub CatAddOwnership()
    On Error GoTo Fail
    Dim serial As String, dcn As String, mk As String, dmk As String
    serial = NV("act_Serial"): dcn = NV("act_DCN")
    mk = NV("act_MakeCode"): dmk = NV("act_DealerMakeCode")

    If Len(serial) = 0 Or Len(dcn) = 0 Then SetResult "Serial Number and DCN are required.": Exit Sub
    If Len(mk) = 0 And Len(dmk) = 0 Then SetResult "Provide Make Code or Dealer Make Code.": Exit Sub

    If MsgBox("Add / update ownership for serial " & serial & ", DCN " & dcn & "?", _
              vbQuestion + vbYesNo, "Confirm Add / Update") <> vbYes Then Exit Sub

    Dim status As Long, txt As String
    txt = CatAddUpdate(serial, dcn, mk, dmk, LCase$(NV("act_OwnType")), NV("act_Model"), _
                       NV("act_ModelYear"), NV("act_PFCode"), NV("act_PFName"), _
                       NV("act_BaseName"), NV("act_CustomName"), status)

    If status = 200 Or status = 201 Then
        SetResult "OK (" & status & "). Record status: " & FieldOf(txt, "status")
    Else
        SetResult "FAILED " & status & ": " & ProxyError(txt, status)
    End If
    Exit Sub
Fail:
    SetResult "ERROR: " & Err.Description
End Sub

Public Sub CatExpireOwnership()
    On Error GoTo Fail
    Dim serial As String, dcn As String, mk As String, dmk As String
    serial = NV("act_Serial"): dcn = NV("act_DCN")
    mk = NV("act_MakeCode"): dmk = NV("act_DealerMakeCode")

    If Len(serial) = 0 Or Len(dcn) = 0 Then SetResult "Serial Number and DCN are required.": Exit Sub
    If Len(mk) = 0 And Len(dmk) = 0 Then SetResult "Provide Make Code or Dealer Make Code.": Exit Sub

    If MsgBox("EXPIRE the ownership record for serial " & serial & ", DCN " & dcn & "?" & _
              vbCrLf & vbCrLf & "This removes the ownership record.", _
              vbExclamation + vbYesNo + vbDefaultButton2, "Confirm Expire") <> vbYes Then Exit Sub

    Dim status As Long, txt As String
    txt = CatExpire(serial, dcn, mk, dmk, status)
    If status = 204 Or status = 200 Then
        SetResult "OK - ownership expired (" & status & ")."
    Else
        SetResult "FAILED " & status & ": " & ProxyError(txt, status)
    End If
    Exit Sub
Fail:
    SetResult "ERROR: " & Err.Description
End Sub

Public Sub CatTransferDecision()
    On Error GoTo Fail
    Dim serial As String, mk As String, dmk As String, st As String, reason As String
    serial = NV("act_Serial"): mk = NV("act_MakeCode"): dmk = NV("act_DealerMakeCode")
    st = UCase$(NV("act_Status")): reason = NV("act_Reason")

    If Len(serial) = 0 Then SetResult "Serial Number is required.": Exit Sub
    If Len(mk) = 0 And Len(dmk) = 0 Then SetResult "Provide Make Code or Dealer Make Code.": Exit Sub
    If st <> "APPROVED" And st <> "REJECTED" Then SetResult "Status must be APPROVED or REJECTED.": Exit Sub
    If st = "REJECTED" And Len(reason) = 0 Then SetResult "A reason is required to reject.": Exit Sub

    If MsgBox(st & " the pending transfer for serial " & serial & "?", _
              vbExclamation + vbYesNo + vbDefaultButton2, "Confirm Transfer") <> vbYes Then Exit Sub

    Dim status As Long, txt As String
    txt = CatTransfer(serial, mk, dmk, st, reason, status)
    If status = 204 Or status = 200 Then
        SetResult "OK - transfer " & st & " (" & status & ")."
    Else
        SetResult "FAILED " & status & ": " & ProxyError(txt, status)
    End If
    Exit Sub
Fail:
    SetResult "ERROR: " & Err.Description
End Sub

Public Sub CatCheckOwnership()
    On Error GoTo Fail
    Dim serial As String: serial = NV("act_Serial")
    If Len(serial) = 0 Then SetResult "Enter a Serial Number to check.": Exit Sub

    Dim txt As String: txt = CatSearch(serial, NV("act_DCN"))   ' read-only
    Dim recs As Object: Set recs = OwnershipRecords(txt)

    If recs Is Nothing Then SetResult "Unexpected response.": Exit Sub
    If recs.Count = 0 Then SetResult "NOT IN CCAT - no ownership records for " & serial: Exit Sub

    Dim msg As String, k As Long, vals As Variant
    msg = recs.Count & " record(s):"
    For k = 1 To recs.Count
        vals = RecordValues(recs(k))                ' 6=DealerCode 7=DealerName 12=Type 13=Status
        msg = msg & vbCrLf & " - " & vals(7) & " (" & vals(6) & ")  " & _
              vals(12) & " / " & vals(13)
    Next k
    SetResult msg
    Exit Sub
Fail:
    SetResult "ERROR: " & Err.Description
End Sub

'==============================================================================
' Read inputs / write result
'==============================================================================
Private Function NV(ByVal nm As String) As String
    On Error Resume Next
    NV = Trim$(CStr(ThisWorkbook.Names(nm).RefersToRange.Value))
End Function

Private Sub SetResult(ByVal msg As String)
    Dim r As Range
    On Error Resume Next
    Set r = ThisWorkbook.Names("act_Result").RefersToRange
    On Error GoTo 0
    If r Is Nothing Then MsgBox msg, vbInformation Else r.Value = Format$(Now, "hh:nn:ss") & "  " & msg
End Sub

'==============================================================================
' One-time setup: build the Actions sheet with inputs and buttons
'==============================================================================
Public Sub CatSetupActionsSheet()
    Dim ws As Worksheet, shp As Shape
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ACTIONS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = ACTIONS_SHEET
    Else
        ws.Cells.Clear
        For Each shp In ws.Shapes: shp.Delete: Next shp
    End If

    ws.Columns("A").ColumnWidth = 24
    ws.Columns("B").ColumnWidth = 36

    PutTitle ws, 1, "Cat Asset Management - Actions (via proxy)"
    ws.Range("A2").Value = "Fill in fields, then click a button. Credentials stay in the proxy; this workbook holds only ProxyUrl + FunctionKey (Config sheet)."
    ws.Range("A3").Value = "Field colour: dark = always required; medium = required for a NEW record; light = optional. Hover a label for details."
    ws.Range("A3").Font.Italic = True

    PutHeader ws, 4, "Asset Identifier (shared by all actions)"
    PutField ws, 5, "Serial Number", "act_Serial", _
        "REQUIRED. Asset serial number (exact match).", 0
    PutField ws, 6, "Make Code", "act_MakeCode", _
        "REQUIRED (this OR Dealer Make Code, not both). Caterpillar manufacturer code, e.g. CW1.", 0
    PutField ws, 7, "Dealer Make Code", "act_DealerMakeCode", _
        "REQUIRED (this OR Make Code, not both). Dealer-specific make code, usually 2 chars, e.g. CW.", 0
    PutField ws, 8, "DCN", "act_DCN", _
        "REQUIRED for Add/Update and Expire. Dealer Customer Number. Not used by Transfer.", 0
    ws.Range("A9").Value = "Provide Make Code OR Dealer Make Code. (DCN is not used for Transfer.)"
    ws.Range("A9").Font.Italic = True

    PutHeader ws, 11, "Add / Update Ownership"
    PutField ws, 12, "Ownership Type", "act_OwnType", _
        "Required for a NEW record. One of: owned, rental, leased, sold, inventory, unknown.", 1
    AddList ws, "act_OwnType", "owned,rental,leased,sold,inventory,unknown"
    PutField ws, 13, "Model", "act_Model", _
        "Required for a NEW record. Asset model, e.g. 980H. Max 65 characters.", 1
    PutField ws, 14, "Model Year", "act_ModelYear", _
        "Required for a NEW record. 4-digit year of manufacture, e.g. 2006.", 1
    PutField ws, 15, "Product Family Code", "act_PFCode", _
        "Optional. Cat product family code, e.g. MDWL. Max 50 characters.", 2
    PutField ws, 16, "Product Family Name", "act_PFName", _
        "Optional. Product family name, e.g. MEDIUM WHEEL LOADER. Max 50 characters.", 2
    PutField ws, 17, "Base Asset Name", "act_BaseName", _
        "Optional. Canonical asset name set by the dealer, e.g. ABC123. Max 60 characters.", 2
    PutField ws, 18, "Custom Asset Name", "act_CustomName", _
        "Optional. Custom asset name, e.g. Excavator #1. Shown in preference to Base Asset Name. Max 60 characters.", 2
    ws.Range("A19").Value = "New records require Ownership Type, Model, Model Year."
    ws.Range("A19").Font.Italic = True
    AddButton ws, "Add / Update", "CatAddOwnership", 12

    PutHeader ws, 21, "Expire Ownership"
    ws.Range("A22").Value = "Uses the shared identifier above (Serial, Make, DCN)."
    AddButton ws, "Expire", "CatExpireOwnership", 22

    PutHeader ws, 24, "Transfer Decision (approve / reject a pending request)"
    PutField ws, 25, "Status", "act_Status", _
        "Transfer only. APPROVED or REJECTED.", 0
    AddList ws, "act_Status", "APPROVED,REJECTED"
    PutField ws, 26, "Reason (required if REJECTED)", "act_Reason", _
        "Required only when Status is REJECTED; optional otherwise.", 2
    AddButton ws, "Approve / Reject", "CatTransferDecision", 25

    PutHeader ws, 28, "Check Ownership (read-only)"
    AddButton ws, "Check Ownership", "CatCheckOwnership", 29

    PutHeader ws, 31, "Result"
    ws.Range("A32").Value = "Last result:"
    With ws.Range("B32")
        .Name = "act_Result"
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    ws.Rows(32).RowHeight = 90

    ws.Range("A34").Value = "Note: write buttons return 'writes_disabled' unless the proxy has CAT_ENABLE_WRITES=true."
    ws.Range("A34").Font.Italic = True

    ws.Activate
    MsgBox "Actions sheet is ready.", vbInformation, "Cat Asset Management (proxy)"
End Sub

'==============================================================================
' Setup helpers
'==============================================================================
Private Sub PutTitle(ByVal ws As Worksheet, ByVal r As Long, ByVal t As String)
    With ws.Cells(r, 1)
        .Value = t
        .Font.Size = 14
        .Font.Bold = True
    End With
End Sub

Private Sub PutHeader(ByVal ws As Worksheet, ByVal r As Long, ByVal t As String)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 2))
        .Merge
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = vbWhite
        .Font.Bold = True
    End With
    ws.Cells(r, 1).Value = t
End Sub

' tier: -1 = no colour; 0 = always required (dark); 1 = required for a NEW
' record (medium); 2 = optional (light). note (if given) becomes a hover comment
' on the label cell.
Private Sub PutField(ByVal ws As Worksheet, ByVal r As Long, ByVal labelText As String, _
                     ByVal nm As String, Optional ByVal note As String = "", _
                     Optional ByVal tier As Long = -1)
    With ws.Cells(r, 1)
        .Value = labelText
        If tier >= 0 Then
            Dim bg As Long, fg As Long
            Select Case tier
                Case 0: bg = RGB(31, 78, 121): fg = vbWhite          ' always required
                Case 1: bg = RGB(46, 117, 182): fg = vbWhite         ' required for a NEW record
                Case Else: bg = RGB(189, 215, 238): fg = RGB(31, 31, 31) ' optional
            End Select
            .Interior.Color = bg
            .Font.Color = fg
            .Font.Bold = True
        End If
        If Len(note) > 0 Then
            If Not .Comment Is Nothing Then .Comment.Delete
            .AddComment note
            .Comment.Shape.TextFrame.AutoSize = True
            .Comment.Visible = False
        End If
    End With
    With ws.Cells(r, 2)
        .Name = nm
        .Interior.Color = RGB(255, 255, 204)
        .Borders.LineStyle = xlContinuous
    End With
End Sub

Private Sub AddList(ByVal ws As Worksheet, ByVal nm As String, ByVal listCsv As String)
    On Error Resume Next
    With ThisWorkbook.Names(nm).RefersToRange.Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:=listCsv
    End With
End Sub

Private Sub AddButton(ByVal ws As Worksheet, ByVal caption As String, _
                      ByVal macroName As String, ByVal anchorRow As Long)
    Dim l As Double, t As Double
    l = ws.Range("D1").Left
    t = ws.Cells(anchorRow, 1).Top
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, l, t, 130, 26)
    btn.OnAction = macroName
    btn.Fill.ForeColor.RGB = RGB(31, 78, 121)
    btn.Line.Visible = msoFalse
    With btn.TextFrame2.TextRange
        .Text = caption
        .Font.Fill.ForeColor.RGB = vbWhite
        .Font.Bold = msoTrue
    End With
End Sub
