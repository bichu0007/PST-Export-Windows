@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Run "Open PST Export.bat" once first so the app can install its dependencies.
  pause
  exit /b 1
)
"%~dp0.venv\Scripts\python.exe" "%~dp0main.py" %*
set "EXITCODE=%ERRORLEVEL%"
echo.
if not "%EXITCODE%"=="0" echo PST Export CLI exited with code %EXITCODE%.
pause
exit /b %EXITCODE%
