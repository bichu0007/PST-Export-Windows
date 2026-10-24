@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title PST Export Toolkit - Setup

echo.
echo ============================================================
echo                 PST EXPORT TOOLKIT
echo            Windows desktop application
echo ============================================================
echo.

REM Locate Python 3.10+.
set "PYTHON_CMD="
where py >nul 2>nul
if not errorlevel 1 set "PYTHON_CMD=py -3"
if not defined PYTHON_CMD (
    where python >nul 2>nul
    if not errorlevel 1 set "PYTHON_CMD=python"
)
if not defined PYTHON_CMD goto :no_python

%PYTHON_CMD% -c "import sys; raise SystemExit(0 if sys.version_info >= (3,10) else 1)" >nul 2>nul
if errorlevel 1 goto :old_python

if not exist ".venv\Scripts\python.exe" (
    echo [1/3] Creating the application's private Python environment...
    %PYTHON_CMD% -m venv ".venv"
    if errorlevel 1 goto :setup_failed
)
set "APP_PY=%~dp0.venv\Scripts\python.exe"
set "APP_PYW=%~dp0.venv\Scripts\pythonw.exe"

"%APP_PY%" -c "import rich, questionary, colorama, tkinter" >nul 2>nul
if errorlevel 1 (
    echo [2/3] Installing application dependencies. This is usually needed only once...
    "%APP_PY%" -m pip install --upgrade pip wheel
    if errorlevel 1 goto :setup_failed
    "%APP_PY%" -m pip install -r requirements.txt
    if errorlevel 1 goto :setup_failed
) else (
    echo [2/3] Standard dependencies are already installed.
)

"%APP_PY%" -c "import tkinter" >nul 2>nul
if errorlevel 1 goto :tk_missing

if not exist ".pypff-install-attempted" (
    echo. > ".pypff-install-attempted"
    "%APP_PY%" -c "import pypff" >nul 2>nul
    if errorlevel 1 (
        echo [3/3] Checking for a prebuilt native PST reader...
        "%APP_PY%" -m pip install --only-binary=:all: libpff-python
        if errorlevel 1 echo [!] Native reader unavailable. The app can use Microsoft Outlook if installed.
    ) else (
        echo [3/3] Native PST reader is already installed.
    )
) else (
    echo [3/3] Native PST reader setup was already attempted.
)

if not exist "%APP_PYW%" goto :pythonw_missing
if not exist "%~dp0gui.py" goto :gui_missing

echo.
echo Opening the PST Export Toolkit desktop app...
start "" "%APP_PYW%" "%~dp0gui.py"
exit /b 0

:no_python
echo [ERROR] Python 3.10 or newer was not found.
echo Install Python from https://www.python.org/downloads/windows/ and enable the Python PATH option.
goto :wait_fail

:old_python
echo [ERROR] Python 3.10 or newer is required.
echo Install a current Python release from https://www.python.org/downloads/windows/.
goto :wait_fail

:pythonw_missing
echo [ERROR] pythonw.exe was not found in the virtual environment.
goto :wait_fail

:tk_missing
echo [ERROR] This Python installation does not include Tkinter. Install the full Python Windows installer with Tcl/Tk support.
goto :wait_fail

:gui_missing
echo [ERROR] gui.py was not found. Re-extract the complete application folder.
goto :wait_fail

:setup_failed
echo [ERROR] Initial setup failed. Check your internet connection and Python installation.
goto :wait_fail

:wait_fail
echo.
pause
exit /b 1
