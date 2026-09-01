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

' --- read-only lookups -------------------------------------------------------
Public Sub CatBatchLookup_Ribbon(control As IRibbonControl)
    ' Module-qualified deliberately: the module CatBatchLookup also contains a
    ' Sub CatBatchLookup, and in a compile-time reference the MODULE name wins -
    ' "Expected variable or procedure, not module".
    CatBatchLookup.CatBatchLookup
End Sub

Public Sub CatBatchLookupDCN_Ribbon(control As IRibbonControl)
    CatBatchLookupDCN
End Sub

' --- build the operation sheets ----------------------------------------------
Public Sub CatBuildAddUpdateSheet_Ribbon(control As IRibbonControl)
    CatBuildAddUpdateSheet
End Sub

Public Sub CatBuildExpireSheet_Ribbon(control As IRibbonControl)
    CatBuildExpireSheet
End Sub

Public Sub CatBuildTransferSheet_Ribbon(control As IRibbonControl)
    CatBuildTransferSheet
End Sub

' --- act on whatever sheet is in front of you --------------------------------
Public Sub CatValidateSheet_Ribbon(control As IRibbonControl)
    CatValidateSheet
End Sub

Public Sub CatRunSheet_Ribbon(control As IRibbonControl)
    CatRunSheet
End Sub

' --- setup -------------------------------------------------------------------
Public Sub CatSettings_Ribbon(control As IRibbonControl)
    CatSettings
End Sub

'==============================================================================
' Per-user settings
'==============================================================================
Public Sub CatSettings()
    Dim url As String, key As String, party As String

    ' InputBox is application-modal: while it is open Excel refuses every click
    ' on the grid and just dings, so you cannot go and copy a value mid-prompt.
    ' Application.InputBox with Type:=8 is the exception - the range picker is
    ' built to let you select cells while it is showing. Offer that first, since
    ' these values usually already live on a sheet (an old Config sheet, say).
    If MsgBox("Import the settings from a worksheet range?" & vbCrLf & vbCrLf & _
              "Yes - point at a two-column range: labels ProxyUrl / FunctionKey /" & vbCrLf & _
              "      PartyNumber on the left, values on the right." & vbCrLf & vbCrLf & _
              "No  - type or paste each value into a prompt (copy it BEFORE you" & vbCrLf & _
              "      start, you cannot reach the sheet once a prompt is open).", _
              vbQuestion + vbYesNo, "Cat Asset Tools - Settings") = vbYes Then
        If ImportSettingsFromRange() Then Exit Sub
        ' cancelled the picker - fall through to the prompts
    End If

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

    ReportSettings "Settings saved for " & Environ$("USERNAME") & "."
End Sub

' Reads ProxyUrl / FunctionKey / PartyNumber from a two-column selection.
' Returns True if anything was imported.
Private Function ImportSettingsFromRange() As Boolean
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox( _
        Prompt:="Select the two-column settings range" & vbCrLf & _
                "(label on the left, value on the right):", _
        Title:="Cat Asset Tools - Import Settings", Type:=8)
    On Error GoTo 0
    If rng Is Nothing Then Exit Function                  ' cancelled

    ' Selecting a whole column would otherwise be a million cells
    Set rng = Application.Intersect(rng, rng.Worksheet.UsedRange)
    If rng Is Nothing Then
        MsgBox "That selection has no data.", vbExclamation, "Cat Asset Tools"
        Exit Function
    End If

    Dim cell As Range, lbl As String, v As String, n As Long
    For Each cell In rng.Columns(1).Cells
        lbl = LCase$(Replace(Trim$(CStr(cell.Value)), " ", ""))
        v = Trim$(CStr(cell.Offset(0, 1).Value))
        Select Case lbl
            Case "proxyurl"
                SaveSetting REG_APP, REG_SECTION, "ProxyUrl", v: n = n + 1
            Case "functionkey"
                SaveSetting REG_APP, REG_SECTION, "FunctionKey", v: n = n + 1
            Case "partynumber"
                SaveSetting REG_APP, REG_SECTION, "PartyNumber", v: n = n + 1
        End Select
    Next cell

    If n = 0 Then
        MsgBox "No recognised labels in that range." & vbCrLf & vbCrLf & _
               "Expected ProxyUrl, FunctionKey or PartyNumber in the left column.", _
               vbExclamation, "Cat Asset Tools"
        Exit Function
    End If

    ImportSettingsFromRange = True
    ReportSettings n & " setting(s) imported for " & Environ$("USERNAME") & "."
End Function

' Shared confirmation - never echoes the key itself, only whether it is set.
Private Sub ReportSettings(ByVal headline As String)
    MsgBox headline & vbCrLf & vbCrLf & _
           "Proxy URL : " & GetSetting(REG_APP, REG_SECTION, "ProxyUrl", "(none)") & vbCrLf & _
           "Key       : " & IIf(Len(GetSetting(REG_APP, REG_SECTION, "FunctionKey", "")) > 0, _
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
