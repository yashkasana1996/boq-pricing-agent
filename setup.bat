@echo off
REM ============================================================
REM  BOQ Agent - one-time setup (Windows). Just double-click me.
REM ============================================================
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo.
  echo Python is not installed or not on PATH.
  echo Install it from https://www.python.org/downloads/
  echo IMPORTANT: tick "Add python.exe to PATH" on the first installer screen.
  echo Then double-click this file again.
  pause
  exit /b 1
)

echo Creating isolated environment (.venv)...
python -m venv .venv
call .venv\Scripts\activate.bat

echo Installing dependencies (needs internet, takes a few minutes)...
pip install --quiet -r requirements.txt

echo Running the full self-test against the sample tender files...
python -m pytest tests -q

echo.
echo ============================================================
echo  If the line above says "8 passed", setup is DONE.
echo  Next: double-click run-app.bat to start the BOQ Agent.
echo ============================================================
pause
