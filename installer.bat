@echo off
setlocal

set SCRIPT_NAME=start.vbs
set SCRIPT_PATH=%~dp0%SCRIPT_NAME%
set STARTUP_FOLDER=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
set LINK_NAME=%STARTUP_FOLDER%\start.vbs.lnk

REM Usuwamy stary skrót jeśli istnieje
if exist "%LINK_NAME%" del "%LINK_NAME%"

REM Tworzymy skrót za pomocą PowerShell
powershell -command "$WshShell = New-Object -ComObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%LINK_NAME%'); $Shortcut.TargetPath = '%SCRIPT_PATH%'; $Shortcut.WorkingDirectory = '%~dp0'; $Shortcut.Save()"

echo Skrót do %SCRIPT_NAME% utworzony w autostarcie.
pause
