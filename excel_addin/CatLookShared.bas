Attribute VB_Name = "CatLookShared"
'==============================================================================
' Cat Asset Management V2 - what the two lookup windows share
'
' frmCatQuickLook (by serial) and frmCatDcnLook (by DCN) both have a Paste at
' Selection button, and both must paste the SAME columns into the same sheet.
' That rule lives here once, so changing how many columns land on Cat Expire
' changes it for both windows at the same time - which is the only way it can
' stay a rule rather than a coincidence.
'
' This is deliberately ALL that is shared. The rest of the two form modules is
' about 90% identical, and it stays duplicated on purpose:
'
'   - the shared version would have to reach into the form's controls, so it
'     would take the form as a late-bound Object. Every control name becomes a
'     string resolved at runtime, so a typo that the compiler catches today
'     would instead fail in front of a user.
'   - it would mean re-pasting both forms' code into the VBE by hand, on a
'     form that took a while to get right.
'
' Two stable forms are not worth either. If a third window ever appears, that
' is the moment to reconsider.
'
' NAMING: the two functions keep their QuickLook* names even though the module
' that used to hold them is now CatSerialLook. Renaming them would mean editing
' both forms' pasted code for no behaviour change, and the working form is not
' worth disturbing for a nicer identifier.
'
' Depends on CatAssetLookup (HeaderArray) and CatBatchOps (the SH_* constants).
'==============================================================================
Option Explicit

' How many recent lookups each window remembers. One number, change it here.
Public Const MRU_MAX As Long = 25

'==============================================================================
' PUBLIC: recent lookups
'
' The dropdown behind the serial / DCN box. Kept here rather than in either
' form so the forms stay thin - and so the two lists cannot accidentally share
' storage, which they must not: a DCN offered as a serial suggestion is noise.
'
' Stored per user in HKCU alongside the proxy settings. Nothing about what you
' looked up belongs inside the add-in file that gets handed around.
'==============================================================================

' Remember a lookup. Newest first, no duplicates, capped at MRU_MAX.
'
' CALL THIS ONLY WHEN THE LOOKUP FOUND SOMETHING. A typo that returned nothing
' is not something anyone wants offered back to them tomorrow, and a list that
' fills up with dead serials stops being worth opening.
Public Sub MruAdd(ByVal kind As String, ByVal value As String)
    On Error Resume Next

    value = Trim$(value)
    If Len(value) = 0 Then Exit Sub

    Dim cur As Variant: cur = MruList(kind)
    Dim out As String, n As Long, i As Long

    out = value                       ' newest always leads
    n = 1

    If IsArray(cur) Then
        For i = LBound(cur) To UBound(cur)
            If n >= MRU_MAX Then Exit For
            ' StrComp text-mode, so "kxh10658" does not join "KXH10658".
            If Len(cur(i)) > 0 And StrComp(CStr(cur(i)), value, vbTextCompare) <> 0 Then
                out = out & "|" & cur(i)
                n = n + 1
            End If
        Next i
    End If

    SaveSetting REG_APP, REG_SECTION, MruKey(kind), out
End Sub

' The remembered values, newest first. Empty array when there are none.
Public Function MruList(ByVal kind As String) As Variant
    On Error Resume Next
    Dim raw As String
    raw = GetSetting(REG_APP, REG_SECTION, MruKey(kind), "")
    If Len(raw) = 0 Then
        MruList = Array()
    Else
        MruList = Split(raw, "|")
    End If
End Function

' Fills a ComboBox from the remembered list.
'
' Takes the control as Object and swallows errors on purpose: if the form still
' has a plain TextBox in that slot - because the control has not been swapped
' yet - .Clear and .AddItem simply do nothing, and the box keeps working as it
' always did. The code is safe to ship ahead of the form change.
Public Sub MruLoad(ByVal box As Object, ByVal kind As String)
    On Error Resume Next

    Dim keep As String
    keep = box.Text                   ' seeding must survive the reload

    box.Clear

    Dim v As Variant: v = MruList(kind)
    Dim i As Long
    If IsArray(v) Then
        For i = LBound(v) To UBound(v)
            If Len(Trim$(CStr(v(i)))) > 0 Then box.AddItem CStr(v(i))
        Next i
    End If

    box.Text = keep
End Sub

Public Sub MruClear(ByVal kind As String)
    On Error Resume Next
    SaveSetting REG_APP, REG_SECTION, MruKey(kind), ""
End Sub

Private Function MruKey(ByVal kind As String) As String
    ' Separate keys, so serials and DCNs can never bleed into one another.
    If UCase$(kind) = "DCN" Then MruKey = "MruDcn" Else MruKey = "MruSerial"
End Function

'==============================================================================
' PUBLIC: send a window's result list to a sheet
'
' You look up a DCN, see 55 assets, and want them in a sheet. Until now that
' meant going and running Batch DCNs over again - a second round trip for data
' already on the screen.
'
' Deliberately the SAME SHAPE as a batch results sheet: the query column, the
' 22 fields, then a Note. That is not cosmetic - Validate and Run refuse any
' sheet with a QuerySerial / QueryDCN column, so a results sheet cannot be
' mistaken for an operation sheet and fired off. Matching the shape inherits
' that refusal.
'==============================================================================
Public Function LookupListToSheet(ByVal rows As Collection, ByVal kind As String, _
                                  ByVal query As String) As String
    If rows Is Nothing Then Exit Function
    If rows.Count = 0 Then Exit Function

    Dim wb As Workbook: Set wb = TargetBook()
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))

    ' Sheet names cap at 31 characters and cannot hold : \ / ? * [ ].
    Dim nm As String
    nm = Left$("Cat " & kind & " " & CleanSheetName(query), 31)
    On Error Resume Next
    ws.Name = nm
    On Error GoTo 0

    Dim h As Variant: h = HeaderArray()
    Dim nFields As Long: nFields = UBound(h) + 1

    ws.Cells(1, 1).Value = "Query" & kind
    Dim c As Long
    For c = 0 To nFields - 1
        ws.Cells(1, 2 + c).Value = h(c)
    Next c
    ws.Cells(1, nFields + 2).Value = "Note"

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, nFields + 2))
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = vbWhite
    End With

    Dim i As Long, v As Variant
    For i = 1 To rows.Count
        v = rows(i)
        ws.Cells(1 + i, 1).NumberFormat = "@"
        ws.Cells(1 + i, 1).Value = query
        For c = 0 To nFields - 1
            ws.Cells(1 + i, 2 + c).NumberFormat = "@"
            ws.Cells(1 + i, 2 + c).Value = CStr(v(c))
        Next c
    Next i

    ws.Columns.AutoFit
    ws.Activate
    LookupListToSheet = ws.Name
End Function

Private Function CleanSheetName(ByVal s As String) As String
    Dim bad As Variant, i As Long
    bad = Array(":", "\", "/", "?", "*", "[", "]")
    For i = LBound(bad) To UBound(bad)
        s = Replace$(s, CStr(bad(i)), "-")
    Next i
    CleanSheetName = s
End Function

' How many columns to write into this sheet. Mirrors the drag-select blocks
' documented in CatBatchOps - each action's set is a prefix of the next:
'
'     Cat Transfer      2 columns   (Serial, Make Code)
'     Cat Expire        3 columns   (+ DCN)
'     Cat Add-Update    6 columns   (+ Ownership Type, Model, Model Year)
'     anything else     all of them
Public Function QuickLookPasteWidth(ByVal ws As Worksheet) As Long
    Select Case ws.Name
        Case SH_TRF: QuickLookPasteWidth = 2
        Case SH_EXP: QuickLookPasteWidth = 3
        Case SH_ADD: QuickLookPasteWidth = 6
        Case Else:   QuickLookPasteWidth = UBound(HeaderArray()) + 1
    End Select
End Function

' Write one record's leading `n` values starting at `target`.
' Forced to text: a serial like 00123 or a DCN that looks numeric would
' otherwise be silently converted to a number and stop matching CCAT.
Public Sub QuickLookPasteAt(ByVal target As Range, ByVal vals As Variant, ByVal n As Long)
    Dim i As Long
    For i = 0 To n - 1
        With target.Offset(0, i)
            .NumberFormat = "@"
            .Value = CStr(vals(i))
        End With
    Next i
End Sub
