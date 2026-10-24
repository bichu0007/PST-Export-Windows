@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Create Desktop Shortcut.ps1"
if errorlevel 1 echo [ERROR] Could not create the shortcut.
echo.
pause
