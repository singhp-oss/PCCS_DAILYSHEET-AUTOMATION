@echo off
:: ══════════════════════════════════════════════════════════════════
::  PCCS Master Startup — Windows Task Scheduler "On Logon" ke liye
::  ------------------------------------------------------------------
::  Kaam:
::    [1] Desktop settle hone do
::    [2] Internet aane tak wait
::    [3] Chrome (WhatsApp Web + Gmail) + remote-debug port
::    [4] PCCS_MASTER.py (watchdog observer + tray) — background
::    [5] Daily Sheet (.xlsm) open
::
::  NOTE: purana WHATSAPP_WATCHER.py OBSOLETE hai — PCCS_MASTER.py khud
::        watchdog observer chalata hai. Woh reference hata diya gaya.
::  NOTE: paths ab project-relative (%~dp0). Hardcoded C:\AWB_TOOLS hataya.
::        DAILY sheet ka path config.json se PCCS_MASTER khud kholta hai;
::        yahan sirf optional pre-open ke liye env var hai (khali chhod
::        sakte ho — tab step [5] skip ho jayega).
:: ══════════════════════════════════════════════════════════════════
title PCCS Master Startup
cd /d "%~dp0"

set "PROJECT_DIR=%~dp0"
set "START_BAT=%PROJECT_DIR%START.bat"

:: Optional: agar daily sheet startup pe kholni ho to yahan poora path do.
:: Khali rakho to PCCS_MASTER khud config se handle karega.
set "EXCEL="

set "CH1=C:\Program Files\Google\Chrome\Application\chrome.exe"
set "CH2=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

:: [1] Desktop load hone do
timeout /t 10 /nobreak >nul

:: [2] Internet wait
:WIFI_LOOP
ping -n 1 -w 1000 8.8.8.8 >nul 2>&1
if errorlevel 1 (
    timeout /t 4 /nobreak >nul
    goto WIFI_LOOP
)

:: [3] Chrome — WhatsApp Web + Gmail, remote-debug port
if exist "%CH1%" (
    start "" "%CH1%" --remote-debugging-port=9222 --new-window https://web.whatsapp.com https://mail.google.com
    goto CHROME_OK
)
if exist "%CH2%" (
    start "" "%CH2%" --remote-debugging-port=9222 --new-window https://web.whatsapp.com https://mail.google.com
)
:CHROME_OK
timeout /t 15 /nobreak >nul

:: [4] PCCS engine — silent launcher (watcher + tray)
if exist "%START_BAT%" (
    start /min "" cmd /c "%START_BAT%"
) else (
    where pythonw >nul 2>&1 && start "" pythonw "%PROJECT_DIR%PCCS_MASTER.py" || start "" python "%PROJECT_DIR%PCCS_MASTER.py"
)
timeout /t 3 /nobreak >nul

:: [5] Daily Sheet (optional pre-open)
if defined EXCEL (
    if exist "%EXCEL%" start "" "%EXCEL%"
)

exit
