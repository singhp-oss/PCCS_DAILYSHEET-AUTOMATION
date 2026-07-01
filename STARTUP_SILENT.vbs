' STARTUP_SILENT.vbs
' Startup folder mein yeh file rakho (BAT ki jagah)
' Koi dialog nahi, koi window nahi - sab background mein

Set WshShell = CreateObject("WScript.Shell")

' MASTER_STARTUP.bat ko bilkul silently chalao
' 0 = hidden window, False = wait mat karo
WshShell.Run "cmd /c ""C:\AWB_TOOLS\MASTER_STARTUP.bat""", 0, False

Set WshShell = Nothing
