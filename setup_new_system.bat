@echo off
REM ============================================================
REM  CSU Advisory Alert - FIRST TIME SETUP (New System)
REM  Run this ONCE on any new machine to install dependencies.
REM  After this, use run_test.bat to verify, then setup_task_scheduler.ps1
REM ============================================================
title CSU Advisory Alert - First Time Setup

echo.
echo  ============================================================
echo   CSU Advisory Alert - First Time Setup
echo  ============================================================
echo.

set SCRIPT_DIR=%~dp0

REM Check Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not on PATH.
    echo         Download from: https://python.org/downloads
    pause
    exit /b 1
)

echo [*] Creating virtual environment...
python -m venv "%SCRIPT_DIR%.venv"

echo [*] Installing dependencies...
"%SCRIPT_DIR%.venv\Scripts\pip" install --upgrade pip --quiet
"%SCRIPT_DIR%.venv\Scripts\pip" install -r "%SCRIPT_DIR%requirements.txt" --quiet

echo.
echo [OK] Setup complete!
echo.
echo  NEXT STEPS:
echo  1. Edit config.ini - update email addresses for this system
echo  2. Double-click run_test.bat to verify emails look correct
echo  3. Right-click setup_task_scheduler.ps1 as Admin to schedule 24/7 runs
echo.
pause