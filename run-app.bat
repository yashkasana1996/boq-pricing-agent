@echo off
REM ============================================================
REM  BOQ Agent - start the web app (Windows). Double-click me.
REM  Your browser opens at http://127.0.0.1:8000
REM  Keep this black window open while you work; close it to stop.
REM ============================================================
cd /d "%~dp0"

if not exist .venv\Scripts\activate.bat (
  echo Please run setup.bat first ^(one time only^).
  pause
  exit /b 1
)
call .venv\Scripts\activate.bat
start "" http://127.0.0.1:8000
python -m uvicorn boq.webapp:app --port 8000
pause
