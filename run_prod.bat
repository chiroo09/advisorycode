@echo off
TITLE CSU Advisory Alert - PRODUCTION RUN
COLOR 0A

echo ============================================================
echo   CSU Threat Intelligence Alert - PRODUCTION RUN
echo ============================================================
echo.

set SCRIPT_DIR=%~dp0
set PYTHON=%SCRIPT_DIR%.venv\Scripts\python.exe
set SCRIPT=%SCRIPT_DIR%advisory_alert_runner.py

if not exist "%PYTHON%" (
    echo [ERROR] Virtual environment not found. Run setup_new_system.bat first.
    pause
    exit /b 1
)

"%PYTHON%" "%SCRIPT%"

echo.
echo ============================================================
echo [*] Production cycle completed.
echo ============================================================
pause