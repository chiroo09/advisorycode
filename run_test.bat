@echo off
REM ============================================================
REM  CSU Advisory Alert - TEST MODE
REM  Uses config.test.ini -> emails go to TEST inbox only.
REM  Opens each email as a PREVIEW in Outlook (no auto-send).
REM  Output saved in: output_test\
REM ============================================================
title CSU Advisory Alert - TEST MODE

set SCRIPT_DIR=%~dp0
set PYTHON=%SCRIPT_DIR%.venv\Scripts\python.exe
set SCRIPT=%SCRIPT_DIR%advisory_alert_runner.py
set TEST_CONFIG=%SCRIPT_DIR%config.test.ini

echo.
echo  ============================================================
echo   CSU Advisory Alert - TEST MODE
echo   Emails: TEST INBOX only (see config.test.ini [email])
echo   Send:   OFF - opens preview in Outlook for review
echo   Data:   test_advisory.xlsx (run run_generate_test.bat first)
echo  ============================================================
echo.

if not exist "%PYTHON%" (
    echo [ERROR] Python venv not found. Run setup_new_system.bat first.
    pause & exit /b 1
)
if not exist "%TEST_CONFIG%" (
    echo [ERROR] config.test.ini not found in: %SCRIPT_DIR%
    pause & exit /b 1
)
if not exist "%SCRIPT_DIR%test_advisory.xlsx" (
    echo [WARN] test_advisory.xlsx not found.
    echo        Run run_generate_test.bat first to create test data.
    echo        Or copy your tracker as test_advisory.xlsx manually.
    echo.
    pause
)

REM Use config.test.ini by temporarily overriding the config path
"%PYTHON%" -c "import advisory_alert_runner as a; a._CONFIG_FILE=r'%TEST_CONFIG%'; a._cfg.read(r'%TEST_CONFIG%', encoding='utf-8'); a.EXCEL_FILE=a._cfg.get('excel','workbook_file',fallback='test_advisory.xlsx'); a.EMAIL_TO=a._cfg.get('email','to'); a.EMAIL_CC=a._cfg.get('email','cc'); a.AUTO_SEND=False; a.ENABLE_EXCEL_AUTO_REFRESH=False; a.MSG_OUTPUT_DIR='output_test'; a.TRACKING_FILE_PREFIX='.test_processed_ids'; a.trigger_excel_macro_refresh=lambda f:None; a.run_cycle()"

echo.
echo  Done. Check Outlook - email previews should be open.
echo  If emails look correct, you are ready for production.
pause