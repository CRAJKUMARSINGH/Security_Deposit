@echo off
setlocal ENABLEDELAYEDEXPANSION

REM Change to this script's directory
cd /d %~dp0

REM Detect Python launcher or python.exe
where py >nul 2>nul
if %ERRORLEVEL%==0 (
  set "PY=py"
) else (
  where python >nul 2>nul
  if %ERRORLEVEL%==0 (
    set "PY=python"
  ) else (
    echo Python not found. Please install Python 3.x from https://www.python.org
    pause
    exit /b 1
  )
)

REM Create virtual environment if missing
if not exist ".venv\Scripts\python.exe" (
  %PY% -m venv .venv
)

set "VENV_PY=.venv\Scripts\python.exe"

REM Upgrade pip and install dependencies
"%VENV_PY%" -m pip install --upgrade pip
"%VENV_PY%" -m pip install --upgrade openpyxl pandas

REM Run the updater script
"%VENV_PY%" update_existing_workbooks.py

echo.
echo Completed updating Excel workbooks.
pause

