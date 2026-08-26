@echo off
setlocal enabledelayedexpansion

TITLE CSU Threat Intelligence Alert Automation
COLOR 0A

echo ========================================================
echo   CSU Threat Intelligence Alert Automation Engine
echo ========================================================
echo.

:: Check if .venv already exists
if exist ".venv\Scripts\activate.bat" goto :ACTIVATE

:: Find Python
echo [*] Checking Python installation...
py -3 --version >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo [*] Creating virtual environment using py launcher...
    py -3 -m venv .venv
    goto :INSTALL
)

python --version >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo [*] Creating virtual environment using python...
    python -m venv .venv
    goto :INSTALL
)

echo [ERROR] Python was not found on your computer.
echo Please install Python from https://www.python.org/
pause
exit /b 1

:INSTALL
if not exist ".venv\Scripts\activate.bat" (
    echo [ERROR] Failed to create virtual environment.
    pause
    exit /b 1
)

:ACTIVATE
echo [*] Activating virtual environment...
call .venv\Scripts\activate.bat

echo [*] Checking dependencies...
pip install -r requirements.txt --quiet

echo.
echo ========================================================
echo [*] Running alert automation...
echo ========================================================
echo.

python advisory_alert_runner.py

echo.
echo ========================================================
echo [*] Script finished.
echo ========================================================
pause
