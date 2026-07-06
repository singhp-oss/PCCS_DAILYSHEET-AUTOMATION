@echo off
:: ══════════════════════════════════════════════════════════════════
::  PCCS AWB System — Silent Launcher (no console window)
::  pythonw = Python without console; tray icon hi UI hai.
::  Project-relative: kahin bhi folder rakho, chalega.
:: ══════════════════════════════════════════════════════════════════
cd /d "%~dp0"

:: pythonw try karo (console-less). Na mile to python fallback.
where pythonw >nul 2>&1
if %errorlevel%==0 (
    start "" pythonw "%~dp0PCCS_MASTER.py"
) else (
    start "" python "%~dp0PCCS_MASTER.py"
)
exit
