@echo off
REM ============================================================
REM  CSU Advisory Alert - Generate Test Data
REM  Copies Alert rows from production tracker into test_advisory.xlsx
REM  Run this BEFORE run_test.bat when testing on a new system.
REM ============================================================
title CSU Advisory Alert - Generate Test Data

set SCRIPT_DIR=%~dp0
set PYTHON=%SCRIPT_DIR%.venv\Scripts\python.exe

echo.
echo  ============================================================
echo   Generating test_advisory.xlsx from production tracker...
echo  ============================================================
echo.

if not exist "%PYTHON%" (
    echo [ERROR] Python venv not found. Run setup_new_system.bat first.
    pause & exit /b 1
)

"%PYTHON%" "%SCRIPT_DIR%generate_test_data.py"

echo.
pause