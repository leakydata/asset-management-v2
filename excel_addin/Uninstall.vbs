Option Explicit

' ============================================================================
'  Cat Asset Tools - uninstaller (VBScript)
'
'  Same job as Uninstall.bat. Removes the add-in file and its registration.
'  Leaves the per-user settings (proxy URL, function key, party number) alone
'  - see the note at the end if you want those gone too.
'
'  Double-clicking shows the result in one dialog. For a console log instead:
'      cscript //nologo Uninstall.vbs
' ============================================================================

Const ADDIN = "ASSET_MANAGEMENT_ADDIN.xlam"

Const HKCU = &H80000001
Const EXCEL_OPTIONS = "Software\Microsoft\Office\{VER}\Excel\Options"

Const ICON_INFO = 64
Const ICON_ERROR = 16

Dim fso, reg, gLog
Set fso = CreateObject("Scripting.FileSystemObject")
Set reg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
gLog = ""

Dim DEST, TARGET
DEST = fso.BuildPath(Env("APPDATA"), "Microsoft\AddIns")
TARGET = fso.BuildPath(DEST, ADDIN)

Say "Cat Asset Tools - removing for " & Env("USERNAME")
Say "============================================================"
Say ""

If ExcelIsRunning() Then
    Say " [X] Excel is open. Close it completely and run this again."
    Finish ICON_ERROR
    WScript.Quit 1
End If

' --- registration ------------------------------------------------------------
'  Delete whichever OPEN slot points at us. Excel closes the gap itself on
'  next start, so a hole in the numbering is not a problem.
Dim ver, key
For Each ver In Array("16.0", "15.0", "14.0")
    key = Replace(EXCEL_OPTIONS, "{VER}", ver)
    If KeyExists(key) Then Deregister key, ver
Next

' --- the file ----------------------------------------------------------------
If fso.FileExists(TARGET) Then
    On Error Resume Next
    fso.DeleteFile TARGET, True
    On Error GoTo 0
    If fso.FileExists(TARGET) Then
        Say " [-] Could not delete " & TARGET
    Else
        Say " [+] Deleted " & TARGET
    End If
Else
    Say " [-] " & ADDIN & " was not in " & DEST
End If

Say ""
Say "============================================================"
Say " Done."
Say ""
Say " Your saved settings were left in place, so a reinstall"
Say " picks them straight back up. To clear those too, run:"
Say "   reg delete ""HKCU\Software\VB and VBA Program Settings\CatAssetTools"" /f"
Say "============================================================"
Finish ICON_INFO
WScript.Quit 0

' ============================================================================
'  Deregister <regkey> <version>
'
'  Finds every OPEN slot whose data holds our filename and deletes those
'  values. Only OPEN-named values are read, so a value whose data merely
'  mentions the name cannot cause the wrong one to be removed.
' ============================================================================
Sub Deregister(regkey, ver)
    Dim names, n, data

    names = ValueNames(regkey)
    For Each n In names
        If UCase(Left(n, 4)) = "OPEN" Then
            data = ""
            reg.GetStringValue HKCU, regkey, n, data
            If Not IsNull(data) Then
                If InStr(1, data, ADDIN, vbTextCompare) > 0 Then
                    If reg.DeleteValue(HKCU, regkey, n) <> 0 Then
                        Say " [-] Could not remove " & n & " from Excel " & ver
                    Else
                        Say " [+] Unregistered from Excel " & ver & " (" & n & ")"
                    End If
                End If
            End If
        End If
    Next
End Sub

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

Function IsConsole()
    IsConsole = (UCase(Right(WScript.FullName, 11)) = "CSCRIPT.EXE")
End Function
