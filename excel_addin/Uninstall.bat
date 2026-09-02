@echo off
setlocal EnableExtensions EnableDelayedExpansion

rem ============================================================================
rem  Cat Asset Tools - uninstaller
rem
rem  Removes the add-in file and its registration. Leaves the per-user settings
rem  (proxy URL, function key, party number) alone - see the note at the end if
rem  you want those gone too.
rem ============================================================================

set "ADDIN=ASSET_MANAGEMENT_ADDIN.xlam"
set "DEST=%APPDATA%\Microsoft\AddIns"

echo.
echo  Cat Asset Tools - removing for %USERNAME%
echo  ============================================================
echo.

tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>nul | find /I "EXCEL.EXE" >nul
if not errorlevel 1 (
    echo  [X] Excel is open. Close it completely and run this again.
    echo.
    pause
    exit /b 1
)

rem --- registration ----------------------------------------------------------
rem  Delete whichever OPEN slot points at us. Excel closes the gap itself on
rem  next start, so a hole in the numbering is not a problem.
for %%V in (16.0 15.0 14.0) do (
    set "KEY=HKCU\Software\Microsoft\Office\%%V\Excel\Options"
    reg query "!KEY!" >nul 2>&1
    if not errorlevel 1 call :deregister "!KEY!" %%V
)

rem --- the file --------------------------------------------------------------
if exist "%DEST%\%ADDIN%" (
    del /F /Q "%DEST%\%ADDIN%" >nul 2>&1
    if exist "%DEST%\%ADDIN%" (
        echo  [-] Could not delete %DEST%\%ADDIN%
    ) else (
        echo  [+] Deleted %DEST%\%ADDIN%
    )
) else (
    echo  [-] %ADDIN% was not in %DEST%
)

echo.
echo  ============================================================
echo   Done.
echo.
echo   Your saved settings were left in place, so a reinstall
echo   picks them straight back up. To clear those too, run:
echo     reg delete "HKCU\Software\VB and VBA Program Settings\CatAssetTools" /f
echo  ============================================================
echo.
pause
exit /b 0

rem ============================================================================
rem  :deregister <regkey> <version>
rem
rem  Finds the OPEN slot holding our filename and deletes that one value. Walks
rem  the slots by name rather than parsing `reg query` output, so a value whose
rem  data merely mentions the name cannot cause the wrong one to be removed.
rem ============================================================================
:deregister
set "KEY=%~1"
set "VER=%~2"

for %%S in (OPEN OPEN1 OPEN2 OPEN3 OPEN4 OPEN5 OPEN6 OPEN7 OPEN8 OPEN9 OPEN10) do (
    reg query "%KEY%" /v %%S 2>nul | find /I "%ADDIN%" >nul
    if not errorlevel 1 (
        reg delete "%KEY%" /v %%S /f >nul 2>&1
        if errorlevel 1 (
            echo  [-] Could not remove %%S from Excel %VER%
        ) else (
            echo  [+] Unregistered from Excel %VER% ^(%%S^)
        )
    )
)
goto :eof
