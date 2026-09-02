Attribute VB_Name = "CatSerialLook"
'==============================================================================
' Cat Asset Management V2 - single-serial look up
'
' Look up ONE serial, see every field for every ownership record it has, then
' copy it or drop it straight into whatever cell you have selected.
'
' Batch lookup already handles one serial, but it spends a whole sheet doing
' it. Checking ten machines one at a time left you with ten results sheets.
' This reuses one floating window instead.
'
' The mirror of this is CatDcnLook: this asks "who owns this machine", that
' asks "what does this customer own". They share the paste rules in
' CatLookShared and nothing else.
'
' The window is MODELESS on purpose: it floats, so you can keep clicking cells
' in the workbook while it is open. That is the whole point - "Paste at
' Selection" resolves Application.ActiveCell at the moment you click it, so
' you can sit on Cat Add-Update row 5, look a serial up, and drop it into A5.
'
' SETUP: this module needs the UserForm 'frmCatQuickLook'. See
' frmCatQuickLook_FORM_CODE.txt - a .frm cannot be handed over as text, so the
' form is built by hand once and the code pasted in.
'
' NAMING: the module is CatSerialLook but the form is still frmCatQuickLook.
' Renaming a UserForm means exporting it, editing the .frm and importing it
' back, and there is nothing to gain from doing that to a form that works.
' The name is only ever seen in the VBE.
'
' Depends on CatAssetLookup (CatSearch, OwnershipRecords, RecordValues,
' HeaderArray, CleanId) and CatLookShared (the paste helpers).
'==============================================================================
Option Explicit

'==============================================================================
' PUBLIC: the button
'==============================================================================
Public Sub CatSerialLook()
    ' Seed from the active cell so "select a serial, click Serial Lookup" just
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
