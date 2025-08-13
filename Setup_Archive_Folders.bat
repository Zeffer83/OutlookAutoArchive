@echo off
REM Setup Archive Folders for Outlook Auto Archive
REM Author: Ryan Zeffiretti
REM Version: 1.0.0

echo ========================================
echo   Outlook Auto Archive - Folder Setup
echo ========================================
echo.
echo This script will create the necessary archive folders
echo and labels for all your email accounts in Outlook.
echo.
echo Make sure Outlook is running before continuing.
echo.

set /p choice="Do you want to continue? (Y/N): "
if /i "%choice%" neq "Y" (
    echo Setup cancelled.
    pause
    exit /b
)

echo.
echo Running setup script...
echo.

powershell.exe -ExecutionPolicy Bypass -File "%~dp0Setup_Archive_Folders.ps1"

echo.
echo Setup script completed.
pause
