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
