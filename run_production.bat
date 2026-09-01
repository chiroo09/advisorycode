@echo off
REM ============================================================
REM  CSU Advisory Alert - PRODUCTION RUN
REM  Sends emails automatically as configured in config.ini.
REM  This is what Windows Task Scheduler calls every 30 minutes.
REM  You can also double-click this to trigger a manual run.
REM ============================================================
title CSU Advisory Alert - Running...

set SCRIPT_DIR=%~dp0
set PYTHON=%SCRIPT_DIR%.venv\Scripts\python.exe
set SCRIPT=%SCRIPT_DIR%advisory_alert_runner.py

if not exist "%PYTHON%" (
    echo [ERROR] Python venv not found. Run setup first.
    pause
    exit /b 1
)

"%PYTHON%" "%SCRIPT%"
exit /b %ERRORLEVEL%