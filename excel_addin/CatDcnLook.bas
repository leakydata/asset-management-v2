Attribute VB_Name = "CatDcnLook"
'==============================================================================
' Cat Asset Management V2 - single-DCN look up
'
' The mirror of CatSerialLook. That one asks "who owns this machine"; this one
' asks "what does this customer own". Same window, same buttons, same paste
' behaviour - the only difference is which argument of CatSearch gets filled.
'
' A serial returns two or three ownership records. A DCN routinely returns
' dozens of assets, which is why the form's record list is taller and its
' columns describe the ASSET rather than the DCN - repeating the same DCN and
' customer name down forty rows tells you nothing.
'
' The window is MODELESS on purpose: it floats, so you can keep clicking cells
' in the workbook while it is open. "Paste at Selection" resolves
' Application.ActiveCell at the moment you click it.
'
' SETUP: this module needs the UserForm 'frmCatDcnLook'. See
' frmCatDcnLook_FORM_CODE.txt - clone frmCatQuickLook by exporting it, renaming
' it in the .frm, and importing it back, then paste the code in.
'
' Depends on CatAssetLookup (CatSearch, OwnershipRecords, RecordValues,
' HeaderArray, CleanId) and on CatLookShared for QuickLookPasteWidth /
' QuickLookPasteAt - the paste rules live there once so a change to how many
' columns land on Cat Expire applies to both windows at once.
'==============================================================================
Option Explicit

'==============================================================================
' PUBLIC: the button
'==============================================================================
Public Sub CatDcnLook()
    ' Seed from the active cell so "select a DCN, click DCN Look" just works.
    ' Guarded on length - an active cell holding a paragraph of notes should
    ' not end up in the box.
    Dim seed As String
    On Error Resume Next
    seed = CleanId(CStr(Application.ActiveCell.Value))
    On Error GoTo 0
    If Len(seed) > 30 Then seed = ""

    frmCatDcnLook.SeedDcn seed
    frmCatDcnLook.Show vbModeless
End Sub
