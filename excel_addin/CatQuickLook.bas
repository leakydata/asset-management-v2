Attribute VB_Name = "CatQuickLook"
'==============================================================================
' Cat Asset Management V2 - single-serial Quick Look
'
' The read counterpart to the operation sheets: look up ONE serial, see every
' field for every ownership record it has, then copy it or drop it straight
' into whatever cell you have selected.
'
' Batch lookup already handles one serial, but it spends a whole sheet doing
' it. Checking ten machines one at a time left you with ten results sheets.
' This reuses one floating window instead.
'
' The window is MODELESS on purpose: it floats, so you can keep clicking cells
' in the workbook while it is open. That is the whole point - "Paste at
' Selection" resolves Application.ActiveCell at the moment you click it, so
' you can sit on Cat Add-Update row 5, look a serial up, and drop it into A5.
'
' The paste is CONTEXTUAL - it writes as many columns as the target sheet
' wants, using the same nesting as the drag-select blocks:
'
'     Cat Transfer      2 columns   (Serial, Make Code)
'     Cat Expire        3 columns   (+ DCN)
'     Cat Add-Update    6 columns   (+ Ownership Type, Model, Model Year)
'     anything else     all of them
'
' SETUP: this module needs the UserForm 'frmCatQuickLook'. See
' frmCatQuickLook_FORM_CODE.txt - a .frm cannot be handed over as text, so the
' form is built by hand once (six controls) and the code pasted in.
'
' Depends on CatAssetLookup (CatSearch, OwnershipRecords, RecordValues,
' HeaderArray, CleanId) and CatBatchOps (the SH_* sheet-name constants).
'==============================================================================
Option Explicit

'==============================================================================
' PUBLIC: the button
'==============================================================================
Public Sub CatQuickLook()
    ' Seed from the active cell so "select a serial, click Quick Look" just
    ' works. Guarded on length - an active cell holding a paragraph of notes
    ' should not end up in the box.
    Dim seed As String
    On Error Resume Next
    seed = CleanId(CStr(Application.ActiveCell.Value))
    On Error GoTo 0
    If Len(seed) > 30 Then seed = ""

    frmCatQuickLook.SeedSerial seed
    frmCatQuickLook.Show vbModeless
End Sub

'==============================================================================
' PUBLIC: shared by the form
'==============================================================================

' How many columns to write into this sheet. Mirrors the drag-select blocks
' documented in CatBatchOps - each action's set is a prefix of the next.
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
