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

    ' Doubles as the add-in's startup hook. onLoad fires when the ribbon is
    ' built, which is when the .xlam loads - and an .xlam has no Workbook_Open
    ' we can ship in a .bas, so this is the one place to hang startup work.
    CatInstallContextMenu
End Sub

'==============================================================================
' Right-click a cell > look it up
'
' The most common action in the add-in, one click from where the data already
' is, instead of a trip to the ribbon.
'
' Temporary:=True is what makes this safe: Excel drops the controls when it
' closes, so an add-in that is uninstalled cannot leave a menu entry behind
' pointing at a macro that no longer exists. The tagged sweep before adding
' handles the other case - a reload inside one Excel session, which would
' otherwise stack up duplicates every time.
'==============================================================================
Private Const CTX_TAG As String = "CatAssetToolsCtx"

Public Sub CatInstallContextMenu()
    On Error Resume Next

    CatRemoveContextMenu

    AddCellItem "Cat: Look up this &serial", "CatCtxSerial"
    AddCellItem "Cat: Look up this &DCN", "CatCtxDcn"
End Sub

Public Sub CatRemoveContextMenu()
    On Error Resume Next
    Dim bar As Object, i As Long
    Set bar = Application.CommandBars("Cell")
    For i = bar.Controls.Count To 1 Step -1
        If bar.Controls(i).Tag = CTX_TAG Then bar.Controls(i).Delete
    Next i
End Sub

Private Sub AddCellItem(ByVal caption As String, ByVal macro As String)
    On Error Resume Next
    Dim btn As Object
    Set btn = Application.CommandBars("Cell").Controls.Add(Type:=1, Temporary:=True)
    btn.Caption = caption
    btn.OnAction = macro
    btn.Tag = CTX_TAG
End Sub

' Both windows already seed themselves from the active cell, and right-clicking
' a cell selects it first - so these need do nothing but open the right one.
Public Sub CatCtxSerial()
    CatSerialLook.CatSerialLook
End Sub

Public Sub CatCtxDcn()
    CatDcnLook.CatDcnLook
End Sub

'==============================================================================
' Button callbacks - thin wrappers over the real macros
'==============================================================================

' --- read-only lookups -------------------------------------------------------
Public Sub CatSerialLook_Ribbon(control As IRibbonControl)
    ' Module-qualified for the same reason as CatBatchLookup below: the module
    ' CatSerialLook also contains a Sub CatSerialLook.
    CatSerialLook.CatSerialLook
End Sub

Public Sub CatDcnLook_Ribbon(control As IRibbonControl)
    ' Module-qualified for the same reason - module CatDcnLook contains a
    ' Sub CatDcnLook, and in a compile-time reference the module name wins.
    CatDcnLook.CatDcnLook
End Sub

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

Public Sub CatUndoRun_Ribbon(control As IRibbonControl)
    CatUndoRun
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

Public Sub CatTestConnection_Ribbon(control As IRibbonControl)
    CatTestConnection
End Sub

Public Sub CatOpenAuditFolder_Ribbon(control As IRibbonControl)
    CatAudit.CatOpenAuditFolder
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
'
' Offers the connection test rather than just reporting, because the moment
' someone has just typed a key is the moment a typo is cheapest to find. Saying
' "saved" and leaving them to discover it does not work is how a five-second
' fix becomes a support call.
Private Sub ReportSettings(ByVal headline As String)
    If MsgBox(headline & vbCrLf & vbCrLf & _
              "Proxy URL : " & GetSetting(REG_APP, REG_SECTION, "ProxyUrl", "(none)") & vbCrLf & _
              "Key       : " & IIf(Len(GetSetting(REG_APP, REG_SECTION, "FunctionKey", "")) > 0, _
                                   "set", "NOT SET") & vbCrLf & vbCrLf & _
              "Test the connection now?", _
              vbQuestion + vbYesNo, "Cat Asset Tools") = vbYes Then
        CatTestConnection
    End If
End Sub

'==============================================================================
' Connection test
'
' Answers "is this thing set up correctly" without needing anyone to read a
' status code. Every outcome names what is wrong AND what to do about it,
' because the person running this is usually the person who cannot fix it by
' reading JSON.
'
' The probe is a search for a serial that deliberately does not exist. The
' proxy answers 200 with an empty record list for an unknown serial - verified
' against the live endpoint - so the test proves the URL, the key and the
' network without depending on any particular asset still being in CCAT. A test
' that fails because someone expired the machine it looked for is worse than no
' test at all.
'
' The status codes below are measured, not assumed:
'     unknown serial   -> 200 with ok:true
'     bad or absent key -> 401
'     wrong URL path    -> 404
'     bad host          -> no reply at all, WinHttp raises
'==============================================================================
Public Sub CatTestConnection()
    Const PROBE_SERIAL As String = "ZZ0000PROBE"

    Dim url As String: url = GetSetting(REG_APP, REG_SECTION, "ProxyUrl", "")
    Dim key As String: key = GetSetting(REG_APP, REG_SECTION, "FunctionKey", "")

    ' Checked here rather than letting CatProxyCall raise, so the two most
    ' common first-run states get their own plain answer.
    If Len(url) = 0 Then
        MsgBox "No proxy URL is set." & vbCrLf & vbCrLf & _
               "Click CCAT > Settings and enter it. It looks like" & vbCrLf & _
               "https://<app>.azurewebsites.net/api", _
               vbExclamation, "Cat Asset Tools - Connection test"
        Exit Sub
    End If
    If Len(key) = 0 Then
        MsgBox "No function key is set." & vbCrLf & vbCrLf & _
               "Click CCAT > Settings and paste it in. Treat it like a " & _
               "password - it is stored under your Windows profile only, " & _
               "never inside the add-in file.", _
               vbExclamation, "Cat Asset Tools - Connection test"
        Exit Sub
    End If

    Dim status As Long, txt As String
    On Error GoTo NoReply
    Application.Cursor = xlWait
    txt = CatProxyCall("GET", "search", "serial=" & PROBE_SERIAL, "", status)
    Application.Cursor = xlDefault
    On Error GoTo 0

    Select Case status
        Case 200
            MsgBox "Connection OK." & vbCrLf & vbCrLf & _
                   "Proxy URL : " & url & vbCrLf & _
                   "Key       : accepted" & vbCrLf & _
                   "Response  : HTTP 200" & vbCrLf & vbCrLf & _
                   "The URL, the key and the network are all working. This " & _
                   "searched for a serial that does not exist on purpose, so " & _
                   "it does not depend on any one asset.", _
                   vbInformation, "Cat Asset Tools - Connection test"

        Case 401, 403
            MsgBox "The proxy answered, but rejected the key." & vbCrLf & vbCrLf & _
                   "Proxy URL : " & url & vbCrLf & _
                   "Response  : HTTP " & status & vbCrLf & vbCrLf & _
                   "So the address and the network are fine - it is the " & _
                   "function key. It is wrong, expired, or has a stray space " & _
                   "on the end from being pasted. Click CCAT > Settings and " & _
                   "enter it again.", _
                   vbExclamation, "Cat Asset Tools - Connection test"

        Case 404
            MsgBox "Reached the server, but there is nothing at that address." & vbCrLf & vbCrLf & _
                   "Proxy URL : " & url & vbCrLf & _
                   "Response  : HTTP 404" & vbCrLf & vbCrLf & _
                   "Usually the URL is missing the /api on the end, or has a " & _
                   "typo. It should look like" & vbCrLf & _
                   "https://<app>.azurewebsites.net/api", _
                   vbExclamation, "Cat Asset Tools - Connection test"

        Case 429
            MsgBox "The proxy is rate-limiting requests (HTTP 429)." & vbCrLf & vbCrLf & _
                   "Your URL and key are fine - this is a throttle, not a " & _
                   "fault. Wait a minute and try again.", _
                   vbInformation, "Cat Asset Tools - Connection test"

        Case 500 To 599
            MsgBox "Reached the proxy, but it returned an error." & vbCrLf & vbCrLf & _
                   "Proxy URL : " & url & vbCrLf & _
                   "Response  : HTTP " & status & vbCrLf & vbCrLf & _
                   "Your URL and key look fine - the problem is server-side. " & _
                   "The Azure Function may be restarting. If it keeps " & _
                   "happening, it needs someone with access to the proxy.", _
                   vbExclamation, "Cat Asset Tools - Connection test"

        Case Else
            MsgBox "Unexpected reply from the proxy." & vbCrLf & vbCrLf & _
                   "Proxy URL : " & url & vbCrLf & _
                   "Response  : HTTP " & status & vbCrLf & vbCrLf & _
                   Left$(txt, 300), _
                   vbExclamation, "Cat Asset Tools - Connection test"
    End Select
    Exit Sub

NoReply:
    ' Nothing answered - DNS, network, VPN or a mistyped host. Deliberately
    ' does NOT mention the key: it was never reached, so blaming it sends
    ' people off to re-paste a key that was fine all along.
    Application.Cursor = xlDefault
    MsgBox "Could not reach the proxy at all." & vbCrLf & vbCrLf & _
           "Proxy URL : " & url & vbCrLf & _
           "Error     : " & Err.Description & vbCrLf & vbCrLf & _
           "Nothing answered, so this is the address or the network - not " & _
           "the key. Check the URL is right, and that you are on the network " & _
           "or VPN you need.", _
           vbCritical, "Cat Asset Tools - Connection test"
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
