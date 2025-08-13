@echo off
echo ========================================
echo    Outlook Auto Archive - With Check
echo ========================================
echo.

echo Checking if Outlook is running...
tasklist /FI "IMAGENAME eq OUTLOOK.EXE" 2>NUL | find /I /N "OUTLOOK.EXE">NUL
if "%ERRORLEVEL%"=="0" (
    echo [OK] Outlook is running
    echo.
    echo Starting archive process...
    echo.
    OutlookAutoArchive.exe
    echo.
    echo Script completed. Check the log files for details.
) else (
    echo [ERROR] Outlook is not running!
    echo.
    echo The archive script requires Outlook to be running.
    echo Please start Outlook and try again.
    echo.
    echo You can:
    echo 1. Start Outlook manually
    echo 2. Run this script again
    echo 3. Set up a scheduled task that runs when Outlook starts
    echo.
)

echo.
echo Press any key to exit...
pause >nul
