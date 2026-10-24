@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Open "Open PST Export.bat" once to install the app and create its environment.
  pause
  exit /b 1
)
"%~dp0.venv\Scripts\python.exe" "%~dp0main.py" %*
set "EXITCODE=%ERRORLEVEL%"
if not "%EXITCODE%"=="0" pause
exit /b %EXITCODE%
