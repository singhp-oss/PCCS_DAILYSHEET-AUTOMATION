@echo off
title PCCS Master Startup
cd /d C:\AWB_TOOLS

set START_BAT=C:\AWB_TOOLS\START.bat
set WATCHER=C:\AWB_TOOLS\WHATSAPP_WATCHER.py
set EXCEL=D:\NEW_PC\DAILY_WORKS\PRADEEP_DAILYSHEET.xlsm
set CH1=C:\Program Files\Google\Chrome\Application\chrome.exe
set CH2=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe

:: [1] Desktop load hone do - 10 sec kaafi hai
timeout /t 10 /nobreak >nul

:: [2] PCCS - /min se minimized CMD, koi popup nahi
if exist "%START_BAT%" (
    start /min "" cmd /c "%START_BAT%"
)
timeout /t 3 /nobreak >nul

:: [3] Internet wait
:WIFI_LOOP
ping -n 1 -w 1000 8.8.8.8 >nul 2>&1
if errorlevel 1 (
    timeout /t 4 /nobreak >nul
    goto WIFI_LOOP
)

:: [4] Chrome - single line
if exist "%CH1%" (
    start "" "%CH1%" --remote-debugging-port=9222 --new-window https://web.whatsapp.com https://mail.google.com
    goto CHROME_OK
)
if exist "%CH2%" (
    start "" "%CH2%" --remote-debugging-port=9222 --new-window https://web.whatsapp.com https://mail.google.com
    goto CHROME_OK
)
:CHROME_OK
timeout /t 18 /nobreak >nul

:: [5] Watcher
if exist "%WATCHER%" (
    start "" pythonw "%WATCHER%"
)
timeout /t 2 /nobreak >nul

:: [6] Excel LAST
if exist "%EXCEL%" (
    start "" "%EXCEL%"
)

exit
