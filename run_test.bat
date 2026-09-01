@echo off
REM ============================================================
REM  CSU Advisory Alert - TEST MODE
REM  Opens email preview in Outlook. Does NOT send automatically.
REM  Use this on any new system to verify before going live.
REM ============================================================
title CSU Advisory Alert - TEST MODE

echo.
echo  ============================================================
echo   CSU Advisory Alert - TEST MODE
echo   Emails will OPEN in Outlook for review, NOT auto-sent.
echo  ============================================================
echo.

REM Temporarily override auto_send to false for this test run
REM We do this by passing a one-liner override via python -c
set SCRIPT_DIR=%~dp0
set PYTHON=%SCRIPT_DIR%.venv\Scripts\python.exe
set SCRIPT=%SCRIPT_DIR%advisory_alert_runner.py

if not exist "%PYTHON%" (
    echo [ERROR] Python venv not found at: %PYTHON%
    echo         Run: python -m venv .venv   then   .venv\Scripts\pip install -r requirements.txt
    pause
    exit /b 1
)

if not exist "%SCRIPT_DIR%config.ini" (
    echo [ERROR] config.ini not found in: %SCRIPT_DIR%
    echo         Copy config.ini from the repository and fill in your details.
    pause
    exit /b 1
)

echo [*] Running in TEST mode (auto_send = false - email preview only)...
echo.

"%PYTHON%" -c "import advisory_alert_runner as a; a.AUTO_SEND=False; a.run_cycle()"

echo.
echo  Done. Check Outlook - draft emails should be open for review.
pause