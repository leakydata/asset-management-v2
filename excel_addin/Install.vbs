Option Explicit

' ============================================================================
'  Cat Asset Tools - installer (VBScript)
'
'  Same job as Install.bat, in a script that double-clicks without a console
'  window flashing up. Unzip anywhere, double-click this file. No admin rights
'  needed: everything it touches is under the current user's own profile.
'
'  What it does, and why each step is here:
'
'   1. Copies the .xlam to %APPDATA%\Microsoft\AddIns. That is Excel's
'      per-user add-in folder AND one of its default Trusted Locations, so
'      the macros run without a security prompt. Anywhere else and the user
'      gets "macros have been disabled" and the ribbon never appears.
'
'   2. The copy also drops the Mark-of-the-Web. A file that arrived in a zip
'      from email or a download carries a Zone.Identifier stream that makes
'      Excel refuse to load it. The copy below rewrites the primary stream
'      into a brand new file, so the block is gone the moment it lands.
'      Note this is NOT true of a plain FileSystemObject.CopyFile, which
'      carries alternate data streams across with the file - which is why
'      the copy here goes through a stream rather than CopyFile.
'
'   3. Registers it, which is the step people miss. Copying the file only
'      makes it APPEAR in File > Options > Add-ins, unticked and not loaded.
'      Excel loads what is listed under Excel\Options in the registry, so the
'      installer writes that entry itself.
'
'  Re-running is safe. It overwrites the file and leaves the registration
'  alone if it is already there.
'
'  Double-clicking shows the result in one dialog. For a console log instead:
'      cscript //nologo Install.vbs
' ============================================================================

' The add-in filename. Change here if it is ever renamed - nothing else in
' this script hardcodes it.
Const ADDIN = "ASSET_MANAGEMENT_ADDIN.xlam"

Const HKCU = &H80000001
Const EXCEL_OPTIONS = "Software\Microsoft\Office\{VER}\Excel\Options"

Const adTypeBinary = 1
Const adSaveCreateOverWrite = 2

Const ICON_INFO = 64
Const ICON_ERROR = 16

Dim fso, reg, gLog, gRegistered
Set fso = CreateObject("Scripting.FileSystemObject")
Set reg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
gLog = ""
gRegistered = False

Dim SRC, DEST
SRC = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), ADDIN)
DEST = fso.BuildPath(Env("APPDATA"), "Microsoft\AddIns")

Say "Cat Asset Tools - installing for " & Env("USERNAME")
Say "============================================================"
Say ""

' --- the file has to actually be next to this script -------------------------
If Not fso.FileExists(SRC) Then
    Say " [X] Could not find " & ADDIN & " next to this installer."
    Say ""
    Say "     Unzip the WHOLE folder first, then run Install.vbs from"
    Say "     inside it. Running it straight out of the zip viewer does"
    Say "     not work - Windows extracts the script to a temp folder on"
    Say "     its own and the .xlam is left behind."
    Fail
End If

' --- Excel must be closed ----------------------------------------------------
'  Two reasons, both of which produce a confusing half-install:
'   - a loaded add-in is locked, so the copy fails
'   - Excel rewrites its Options key when it exits, so a registration made
'     while it is running gets overwritten on close
If ExcelIsRunning() Then
    Say " [X] Excel is open. Close it completely and run this again."
    Say ""
    Say "     Check the system tray and Task Manager if you think it is"
    Say "     already closed - a hidden EXCEL.EXE will block the install."
    Fail
End If

' --- 1. copy -----------------------------------------------------------------
If Not fso.FolderExists(DEST) Then
    On Error Resume Next
    fso.CreateFolder DEST
    On Error GoTo 0
End If
If Not fso.FolderExists(DEST) Then
    Say " [X] Could not create " & DEST
    Fail
End If

If Not CopyWithoutStreams(SRC, fso.BuildPath(DEST, ADDIN)) Then
    Say " [X] Could not copy the add-in to " & DEST
    Fail
End If
Say " [1/2] Copied to " & DEST

' --- 2. register -------------------------------------------------------------
'  Excel 16.0 is 2016 / 2019 / 2021 / 365, 15.0 is 2013, 14.0 is 2010. Only
'  versions actually installed have an Excel\Options key, so the rest are
'  skipped silently rather than guessed at.
Dim ver, key
For Each ver In Array("16.0", "15.0", "14.0")
    key = Replace(EXCEL_OPTIONS, "{VER}", ver)
    If KeyExists(key) Then Register key, ver
Next

If Not gRegistered Then
    Say " [2/2] Could not find an installed Excel to register with."
    Say ""
    Say "       The add-in is copied and ready. Turn it on by hand:"
    Say "       Excel > File > Options > Add-ins > Manage: Excel"
    Say "       Add-ins > Go... > tick """ & ADDIN & """"
End If

Say ""
Say "============================================================"
Say " Done. Open Excel and look for the CCAT tab."
Say ""
Say " First run only: CCAT > Settings, and enter the proxy URL"
Say " and function key. They are stored per-user and are not in"
Say " this zip."
Say "============================================================"
Finish ICON_INFO
WScript.Quit 0

' ============================================================================
'  Register <regkey> <version>
'
'  Excel keeps its add-in list as values named OPEN, OPEN1, OPEN2 ... under
'  Excel\Options. There is no "add" - you write the next free one. Taking a
'  slot that is already in use would silently unload whatever add-in was
'  there, so the free slot is searched for rather than assumed.
' ============================================================================
Sub Register(regkey, ver)
    Dim names, slot

    names = ValueNames(regkey)

    ' Already listed? Then leave it exactly as it is - a second entry for the
    ' same file makes Excel load it twice and complain about the name.
    If OpenSlotFor(regkey, names, ADDIN) <> "" Then
        Say " [2/2] Already registered with Excel " & ver
        gRegistered = True
        Exit Sub
    End If

    ' First free slot: OPEN, then OPEN1..OPEN20.
    slot = FirstFreeSlot(names)
    If slot = "" Then
        Say " [2/2] Excel " & ver & " already has 21 add-ins listed - no free slot."
        Exit Sub
    End If

    ' The value data includes the quotes: Excel stores "NAME.xlam", and a bare
    ' filename resolves against the AddIns folder we just copied into.
    If reg.SetStringValue(HKCU, regkey, slot, """" & ADDIN & """") <> 0 Then
        Say " [2/2] Copied, but could not write the registry for Excel " & ver & "."
        Say "       Turn it on by hand: File > Options > Add-ins > Manage:"
        Say "       Excel Add-ins > Go... > tick """ & ADDIN & """"
        Exit Sub
    End If

    Say " [2/2] Registered with Excel " & ver & " (" & slot & ")"
    gRegistered = True
End Sub

' First unused name out of OPEN, OPEN1..OPEN20, or "" if every one is taken.
Function FirstFreeSlot(names)
    Dim i, candidate
    FirstFreeSlot = ""
    For i = 0 To 20
        If i = 0 Then
            candidate = "OPEN"
        Else
            candidate = "OPEN" & i
        End If
        If Not InList(names, candidate) Then
            FirstFreeSlot = candidate
            Exit Function
        End If
    Next
End Function

' ============================================================================
'  CopyWithoutStreams <src> <dest>
'
'  ADODB.Stream reads the bytes and writes a brand new file, so anything
'  hanging off the source as an alternate data stream - Zone.Identifier above
'  all - does not come with it. If ADODB is unavailable for any reason we fall
'  back to an ordinary copy and then try to delete the stream by name.
' ============================================================================
Function CopyWithoutStreams(src, dest)
    Dim st
    CopyWithoutStreams = False

    On Error Resume Next
    Set st = CreateObject("ADODB.Stream")
    If Err.Number = 0 Then
        st.Type = adTypeBinary
        st.Open
        st.LoadFromFile src
        st.SaveToFile dest, adSaveCreateOverWrite
        st.Close
    End If

    If Err.Number <> 0 Then
        ' Fall back: ordinary copy, then strip the Mark-of-the-Web by hand.
        Err.Clear
        fso.CopyFile src, dest, True
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Exit Function
        End If
        fso.DeleteFile dest & ":Zone.Identifier", True
        Err.Clear
    End If
    On Error GoTo 0

    CopyWithoutStreams = fso.FileExists(dest)
End Function

' ============================================================================
'  Shared helpers
' ============================================================================

' True if a registry key exists. EnumValues returns 0 for a key that is there
' even when it holds no values, which is exactly the test we want.
Function KeyExists(subkey)
    Dim names, types
    KeyExists = (reg.EnumValues(HKCU, subkey, names, types) = 0)
End Function

' The value names under a key, as an array - empty if the key holds none.
Function ValueNames(subkey)
    Dim names, types
    If reg.EnumValues(HKCU, subkey, names, types) <> 0 Then
        ValueNames = Array()
    ElseIf IsNull(names) Then
        ValueNames = Array()
    Else
        ValueNames = names
    End If
End Function

' The OPEN* value holding the given filename, or "" if it is not listed. Only
' OPEN-named values are read, so a value whose data merely mentions the name
' cannot be mistaken for a registration.
Function OpenSlotFor(subkey, names, filename)
    Dim n, data
    OpenSlotFor = ""
    For Each n In names
        If UCase(Left(n, 4)) = "OPEN" Then
            data = ""
            reg.GetStringValue HKCU, subkey, n, data
            If Not IsNull(data) Then
                If InStr(1, data, filename, vbTextCompare) > 0 Then
                    OpenSlotFor = n
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Function InList(arr, value)
    Dim x
    InList = False
    For Each x In arr
        If UCase(x) = UCase(value) Then
            InList = True
            Exit Function
        End If
    Next
End Function

Function ExcelIsRunning()
    Dim wmi, procs
    Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set procs = wmi.ExecQuery("SELECT Name FROM Win32_Process WHERE Name = 'EXCEL.EXE'")
    ExcelIsRunning = (procs.Count > 0)
End Function

Function Env(name)
    Dim sh
    Set sh = CreateObject("WScript.Shell")
    Env = sh.ExpandEnvironmentStrings("%" & name & "%")
End Function

' Under cscript every line prints as it happens; under wscript they are held
' and shown together at the end, so a double-click gets one dialog rather than
' a dozen.
Sub Say(line)
    If IsConsole() Then
        WScript.Echo line
    Else
        gLog = gLog & line & vbCrLf
    End If
End Sub

Sub Finish(icon)
    If Not IsConsole() Then MsgBox gLog, icon, "Cat Asset Tools"
End Sub

Sub Fail()
    Say ""
    Say "============================================================"
    Say " Install failed - nothing was changed."
    Say "============================================================"
    Finish ICON_ERROR
    WScript.Quit 1
End Sub

Function IsConsole()
    IsConsole = (UCase(Right(WScript.FullName, 11)) = "CSCRIPT.EXE")
End Function
