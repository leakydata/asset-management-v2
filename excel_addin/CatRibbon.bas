Attribute VB_Name = "CatRibbon"
'==============================================================================
' Cat Asset Management V2 - ribbon callbacks + per-user settings
'
' Pair this module with customUI14.xml, injected into the .xlam with the
' Office RibbonX Editor (Insert > Office 2010+ Custom UI Part).
'
' WHY THE WRAPPERS: a ribbon onAction callback must have the signature
'   Sub Name(control As IRibbonControl)
' The existing macros take no arguments. Excel is often forgiving about this,
' but it is not documented to be, and a silent no-op on a button your boss is
' clicking is not a failure mode worth risking. Each wrapper below just calls
' the real macro, which stays callable from Alt+F8, from other code, and from
' a shape's Assign Macro exactly as before.
'
' WHY THE REGISTRY: in an add-in, ThisWorkbook is the hidden .xlam itself, so
' the old Config worksheet would ship inside the file you hand to other people
' - function key and all. GetSetting/SaveSetting keeps each person's proxy URL
' and key in their own HKCU hive instead. Nothing sensitive travels with the
' add-in.
'
' Replace the old Cfg() in CatAssetLookup with:
'
'     Private Function Cfg(ByVal label As String) As String
'         Cfg = GetSetting(REG_APP, REG_SECTION, label, "")
'     End Function
'
' The labels are unchanged - "ProxyUrl", "FunctionKey", "PartyNumber" - so no
' call site needs editing.
'==============================================================================
Option Explicit

Public Const REG_APP     As String = "CatAssetTools"
Public Const REG_SECTION As String = "Proxy"

' Cached so a later version can call ribbonUI.Invalidate to refresh labels
' (e.g. showing whether the key is configured). Nothing uses it yet.
Private ribbonUI As IRibbonUI

'==============================================================================
' Ribbon load
'==============================================================================
Public Sub CatRibbonOnLoad(ribbon As IRibbonUI)
    Set ribbonUI = ribbon
End Sub

'==============================================================================
' Button callbacks - thin wrappers over the real macros
'==============================================================================
Public Sub CatCheckOwnership_Ribbon(control As IRibbonControl)
    CatCheckOwnership
End Sub

Public Sub CatBatchLookup_Ribbon(control As IRibbonControl)
    CatBatchLookup
End Sub

Public Sub CatBatchLookupDCN_Ribbon(control As IRibbonControl)
    CatBatchLookupDCN
End Sub

Public Sub CatAddOwnership_Ribbon(control As IRibbonControl)
    CatAddOwnership
End Sub

Public Sub CatExpireOwnership_Ribbon(control As IRibbonControl)
    CatExpireOwnership
End Sub

Public Sub CatTransferDecision_Ribbon(control As IRibbonControl)
    CatTransferDecision
End Sub

Public Sub CatSetupActionsSheet_Ribbon(control As IRibbonControl)
    CatSetupActionsSheet
End Sub

Public Sub CatBatchAddUpdate_Ribbon(control As IRibbonControl)
    CatBatchAddUpdate
End Sub

Public Sub CatSetupBatchAddUpdateSheet_Ribbon(control As IRibbonControl)
    CatSetupBatchAddUpdateSheet
End Sub

Public Sub CatSettings_Ribbon(control As IRibbonControl)
    CatSettings
End Sub

'==============================================================================
' Per-user settings
'==============================================================================
Public Sub CatSettings()
    Dim url As String, key As String, party As String

    url = InputBox( _
        "Proxy URL" & vbCrLf & vbCrLf & _
        "e.g. https://<app>.azurewebsites.net/api", _
        "Cat Asset Tools - Settings (1 of 3)", _
        GetSetting(REG_APP, REG_SECTION, "ProxyUrl", ""))
    If StrPtr(url) = 0 Then Exit Sub          ' Cancel, not an empty string
    SaveSetting REG_APP, REG_SECTION, "ProxyUrl", Trim$(url)

    key = InputBox( _
        "Function key" & vbCrLf & vbCrLf & _
        "Treat this like a password. It is stored under your Windows " & _
        "profile only - it is never saved into the add-in file.", _
        "Cat Asset Tools - Settings (2 of 3)", _
        GetSetting(REG_APP, REG_SECTION, "FunctionKey", ""))
    If StrPtr(key) = 0 Then Exit Sub
    SaveSetting REG_APP, REG_SECTION, "FunctionKey", Trim$(key)

    party = InputBox( _
        "Party number (optional)" & vbCrLf & vbCrLf & _
        "Leave blank unless an admin tells you otherwise - the proxy " & _
        "supplies the dealer code server-side.", _
        "Cat Asset Tools - Settings (3 of 3)", _
        GetSetting(REG_APP, REG_SECTION, "PartyNumber", ""))
    If StrPtr(party) = 0 Then Exit Sub
    SaveSetting REG_APP, REG_SECTION, "PartyNumber", Trim$(party)

    MsgBox "Settings saved for " & Environ$("USERNAME") & "." & vbCrLf & vbCrLf & _
           "Proxy URL : " & GetSetting(REG_APP, REG_SECTION, "ProxyUrl", "(none)") & vbCrLf & _
           "Key        : " & IIf(Len(GetSetting(REG_APP, REG_SECTION, "FunctionKey", "")) > 0, _
                                 "set", "NOT SET"), _
           vbInformation, "Cat Asset Tools"
End Sub

' Clears this user's saved settings - useful on a shared machine.
Public Sub CatClearSettings()
    If MsgBox("Clear the saved proxy URL and function key for " & _
              Environ$("USERNAME") & "?", vbQuestion + vbYesNo, _
              "Cat Asset Tools") <> vbYes Then Exit Sub
    On Error Resume Next
    DeleteSetting REG_APP, REG_SECTION
    On Error GoTo 0
    MsgBox "Settings cleared.", vbInformation, "Cat Asset Tools"
End Sub
