Attribute VB_Name = "CatActions"
'==============================================================================
' Cat Asset Management V2 - actions (add/update, expire, transfer, check)
'
' These call WRITE endpoints, so they are deliberate macros (never worksheet
' functions) and each asks for confirmation. Run CatSetupActionsSheet once to
' build an "Actions" sheet with input cells and buttons.
'
' Depends on the CatAssetLookup module (CatAddUpdate, CatExpire, CatTransfer,
' CatSearch, RecordValues) and VBA-JSON.
'
' WARNING: add / expire / transfer change real data. Use your test dealer code
' (Config!PartyNumber) and test assets until you intend to touch production.
'==============================================================================
Option Explicit

Private Const ACTIONS_SHEET As String = "Actions"

'==============================================================================
' Action buttons
'==============================================================================
Public Sub CatAddOwnership()
    On Error GoTo Fail
    If CAT_READ_ONLY Then SetResult "Read-only mode: writes are disabled.": Exit Sub
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
        SetResult "FAILED " & status & ": " & ErrText(txt)
    End If
    Exit Sub
Fail:
    SetResult "ERROR: " & Err.Description
End Sub

Public Sub CatExpireOwnership()
    On Error GoTo Fail
    If CAT_READ_ONLY Then SetResult "Read-only mode: writes are disabled.": Exit Sub
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
    If status = 204 Then
        SetResult "OK - ownership expired (204)."
    Else
        SetResult "FAILED " & status & ": " & ErrText(txt)
    End If
    Exit Sub
Fail:
    SetResult "ERROR: " & Err.Description
End Sub

Public Sub CatTransferDecision()
    On Error GoTo Fail
    If CAT_READ_ONLY Then SetResult "Read-only mode: writes are disabled.": Exit Sub
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
    If status = 204 Then
        SetResult "OK - transfer " & st & " (204)."
    Else
        SetResult "FAILED " & status & ": " & ErrText(txt)
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
    Dim j As Object: Set j = JsonConverter.ParseJson(txt)
    Dim recs As Object
    If Not j Is Nothing Then If j.Exists("ownershipRecords") Then Set recs = j("ownershipRecords")

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

Private Function FieldOf(ByVal txt As String, ByVal key As String) As String
    On Error Resume Next
    Dim j As Object: Set j = JsonConverter.ParseJson(txt)
    If Not j Is Nothing Then If j.Exists(key) Then FieldOf = CStr(j(key))
End Function

Private Function ErrText(ByVal txt As String) As String
    On Error Resume Next
    Dim j As Object: Set j = JsonConverter.ParseJson(txt)
    If Not j Is Nothing Then
        If j.Exists("code") Then ErrText = CStr(j("code")) & " "
        If j.Exists("description") Then ErrText = ErrText & CStr(j("description"))
        Dim det As Object
        If j.Exists("details") Then
            Set det = j("details")
            If det.Count >= 1 Then
                If det(1).Exists("message") Then ErrText = ErrText & " (" & CStr(det(1)("message")) & ")"
            End If
        End If
    End If
    If Len(Trim$(ErrText)) = 0 Then ErrText = Left$(txt, 200)
End Function

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

    PutTitle ws, 1, "Cat Asset Management - Actions"
    ws.Range("A2").Value = "Fill in fields, then click a button. partyNumber comes from the Config sheet."
    If CAT_READ_ONLY Then
        With ws.Range("A3")
            .Value = "READ-ONLY MODE: Add/Update, Expire and Transfer are disabled. Only Check Ownership is available."
            .Font.Bold = True
            .Font.Color = RGB(192, 0, 0)
        End With
    End If

    PutHeader ws, 4, "Asset Identifier (shared by all actions)"
    PutField ws, 5, "Serial Number", "act_Serial"
    PutField ws, 6, "Make Code", "act_MakeCode"
    PutField ws, 7, "Dealer Make Code", "act_DealerMakeCode"
    PutField ws, 8, "DCN", "act_DCN"
    ws.Range("A9").Value = "Provide Make Code OR Dealer Make Code. (DCN is not used for Transfer.)"
    ws.Range("A9").Font.Italic = True

    PutHeader ws, 11, "Add / Update Ownership"
    PutField ws, 12, "Ownership Type", "act_OwnType"
    AddList ws, "act_OwnType", "owned,rental,leased,sold,inventory,unknown"
    PutField ws, 13, "Model", "act_Model"
    PutField ws, 14, "Model Year", "act_ModelYear"
    PutField ws, 15, "Product Family Code", "act_PFCode"
    PutField ws, 16, "Product Family Name", "act_PFName"
    PutField ws, 17, "Base Asset Name", "act_BaseName"
    PutField ws, 18, "Custom Asset Name", "act_CustomName"
    ws.Range("A19").Value = "New records require Ownership Type, Model, Model Year."
    ws.Range("A19").Font.Italic = True
    If Not CAT_READ_ONLY Then AddButton ws, "Add / Update", "CatAddOwnership", 12

    PutHeader ws, 21, "Expire Ownership"
    ws.Range("A22").Value = "Uses the shared identifier above (Serial, Make, DCN)."
    If Not CAT_READ_ONLY Then AddButton ws, "Expire", "CatExpireOwnership", 22

    PutHeader ws, 24, "Transfer Decision (approve / reject a pending request)"
    PutField ws, 25, "Status", "act_Status"
    AddList ws, "act_Status", "APPROVED,REJECTED"
    PutField ws, 26, "Reason (required if REJECTED)", "act_Reason"
    If Not CAT_READ_ONLY Then AddButton ws, "Approve / Reject", "CatTransferDecision", 25

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

    ws.Activate
    MsgBox "Actions sheet is ready.", vbInformation, "Cat Asset Management"
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

Private Sub PutField(ByVal ws As Worksheet, ByVal r As Long, ByVal labelText As String, ByVal nm As String)
    ws.Cells(r, 1).Value = labelText
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
