@echo off
setlocal EnableExtensions EnableDelayedExpansion

rem ============================================================================
rem  Cat Asset Tools - installer
rem
rem  Unzip anywhere, double-click this file. No admin rights needed: everything
rem  it touches is under the current user's own profile.
rem
rem  What it does, and why each step is here:
rem
rem   1. Copies the .xlam to %APPDATA%\Microsoft\AddIns. That is Excel's
rem      per-user add-in folder AND one of its default Trusted Locations, so
rem      the macros run without a security prompt. Anywhere else and the user
rem      gets "macros have been disabled" and the ribbon never appears.
rem
rem   2. The copy also drops the Mark-of-the-Web. A file that arrived in a zip
rem      from email or a download carries a Zone.Identifier stream that makes
rem      Excel refuse to load it. CMD's COPY writes only the primary stream, so
rem      the block is gone the moment it lands.
rem
rem   3. Registers it, which is the step people miss. Copying the file only
rem      makes it APPEAR in File > Options > Add-ins, unticked and not loaded.
rem      Excel loads what is listed under Excel\Options in the registry, so the
rem      installer writes that entry itself.
rem
rem  Re-running is safe. It overwrites the file and leaves the registration
rem  alone if it is already there.
rem ============================================================================

rem The add-in filename. Change here if it is ever renamed - nothing else in
rem this script hardcodes it.
set "ADDIN=ASSET_MANAGEMENT_ADDIN.xlam"

set "SRC=%~dp0%ADDIN%"
set "DEST=%APPDATA%\Microsoft\AddIns"

echo.
echo  Cat Asset Tools - installing for %USERNAME%
echo  ============================================================
echo.

rem --- the file has to actually be next to this script -----------------------
if not exist "%SRC%" (
    echo  [X] Could not find %ADDIN% next to this installer.
    echo.
    echo      Unzip the WHOLE folder first, then run Install.bat from
    echo      inside it. Running it straight out of the zip viewer does
    echo      not work - Windows extracts the .bat to a temp folder on
    echo      its own and the .xlam is left behind.
    echo.
    goto :fail
)

rem --- Excel must be closed --------------------------------------------------
rem  Two reasons, both of which produce a confusing half-install:
rem   - a loaded add-in is locked, so the copy fails
rem   - Excel rewrites its Options key when it exits, so a registration made
rem     while it is running gets overwritten on close
tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>nul | find /I "EXCEL.EXE" >nul
if not errorlevel 1 (
    echo  [X] Excel is open. Close it completely and run this again.
    echo.
    echo      Check the system tray and Task Manager if you think it is
    echo      already closed - a hidden EXCEL.EXE will block the install.
    echo.
    goto :fail
)

rem --- 1. copy ---------------------------------------------------------------
if not exist "%DEST%" mkdir "%DEST%" 2>nul
if not exist "%DEST%" (
    echo  [X] Could not create %DEST%
    goto :fail
)

copy /Y "%SRC%" "%DEST%\%ADDIN%" >nul
if errorlevel 1 (
    echo  [X] Could not copy the add-in to %DEST%
    goto :fail
)
echo  [1/2] Copied to %DEST%

rem --- 2. register -----------------------------------------------------------
rem  Excel 16.0 is 2016 / 2019 / 2021 / 365, 15.0 is 2013, 14.0 is 2010. Only
rem  versions actually installed have an Excel\Options key, so the rest are
rem  skipped silently rather than guessed at.
set "REGISTERED=0"
for %%V in (16.0 15.0 14.0) do (
    set "KEY=HKCU\Software\Microsoft\Office\%%V\Excel\Options"
    reg query "!KEY!" >nul 2>&1
    if not errorlevel 1 call :register "!KEY!" %%V
)

if "%REGISTERED%"=="0" (
    echo  [2/2] Could not find an installed Excel to register with.
    echo.
    echo        The add-in is copied and ready. Turn it on by hand:
    echo        Excel ^> File ^> Options ^> Add-ins ^> Manage: Excel
    echo        Add-ins ^> Go... ^> tick "%ADDIN%"
    echo.
    goto :done
)

:done
echo.
echo  ============================================================
echo   Done. Open Excel and look for the CCAT tab.
echo.
echo   First run only: CCAT ^> Settings, and enter the proxy URL
echo   and function key. They are stored per-user and are not in
echo   this zip.
echo  ============================================================
echo.
pause
exit /b 0

:fail
echo  ============================================================
echo   Install failed - nothing was changed.
echo  ============================================================
echo.
pause
exit /b 1

rem ============================================================================
rem  :register <regkey> <version>
rem
rem  Excel keeps its add-in list as values named OPEN, OPEN1, OPEN2 ... under
rem  Excel\Options. There is no "add" - you write the next free one. Taking a
rem  slot that is already in use would silently unload whatever add-in was
rem  there, so the free slot is searched for rather than assumed.
rem ============================================================================
:register
set "KEY=%~1"
set "VER=%~2"

rem Already listed? Then leave it exactly as it is - a second entry for the
rem same file makes Excel load it twice and complain about the name.
reg query "%KEY%" 2>nul | find /I "%ADDIN%" >nul
if not errorlevel 1 (
    echo  [2/2] Already registered with Excel %VER%
    set "REGISTERED=1"
    goto :eof
)

rem First free slot: OPEN, then OPEN1..OPEN20.
set "SLOT="
reg query "%KEY%" /v OPEN >nul 2>&1
if errorlevel 1 (
    set "SLOT=OPEN"
) else (
    for /L %%N in (1,1,20) do (
        if not defined SLOT (
            reg query "%KEY%" /v OPEN%%N >nul 2>&1
            if errorlevel 1 set "SLOT=OPEN%%N"
        )
    )
)

if not defined SLOT (
    echo  [2/2] Excel %VER% already has 21 add-ins listed - no free slot.
    goto :eof
)

rem The value data includes the quotes: Excel stores "NAME.xlam", and a bare
rem filename resolves against the AddIns folder we just copied into.
reg add "%KEY%" /v "%SLOT%" /t REG_SZ /d "\"%ADDIN%\"" /f >nul 2>&1
if errorlevel 1 (
    echo  [2/2] Copied, but could not write the registry for Excel %VER%.
    echo        Turn it on by hand: File ^> Options ^> Add-ins ^> Manage:
    echo        Excel Add-ins ^> Go... ^> tick "%ADDIN%"
    goto :eof
)

echo  [2/2] Registered with Excel %VER% (%SLOT%)
set "REGISTERED=1"
goto :eof
