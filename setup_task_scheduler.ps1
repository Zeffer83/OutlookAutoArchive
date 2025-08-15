# Setup Task Scheduler for Outlook Auto Archive
# This script creates a scheduled task to run the archive application automatically
# Version: 2.9.5
# Author: Ryan Zeffiretti
# Description: Task scheduler setup for Outlook Auto Archive
# License: MIT

# ASCII Art Banner
Write-Host ""
Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║                                                              ║" -ForegroundColor Cyan
Write-Host "║   OUTLOOK AUTO ARCHIVE - TASK SCHEDULER                      ║" -ForegroundColor Cyan
Write-Host "║   Automated Email Archiving Setup                            ║" -ForegroundColor Cyan
Write-Host "║                                                              ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
Write-Host "[TARGET] This script will set up automatic archiving for your Outlook emails" -ForegroundColor White
Write-Host "[SCHEDULE] Choose when you want the archiving to run automatically" -ForegroundColor White
Write-Host ""

# Check if running as Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

if (-not $isAdmin) {
    Write-Host "[ERROR] This script requires Administrator privileges to create scheduled tasks." -ForegroundColor Red
    Write-Host ""
    Write-Host "Please:" -ForegroundColor Yellow
    Write-Host "1. Right-click on PowerShell" -ForegroundColor White
    Write-Host "2. Select 'Run as Administrator'" -ForegroundColor White
    Write-Host "3. Navigate to this directory" -ForegroundColor White
    Write-Host "4. Run this script again" -ForegroundColor White
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

Write-Host "[OK] Running with Administrator privileges" -ForegroundColor Green
Write-Host ""

# Define the executable path
$exePath = "C:\Users\$env:USERNAME\OutlookAutoArchive\OutlookAutoArchive.exe"

# Check if the executable exists
if (-not (Test-Path $exePath)) {
    Write-Host "[ERROR] Executable not found at: $exePath" -ForegroundColor Red
    Write-Host "Please make sure the Outlook Auto Archive application is installed." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

Write-Host "[OK] Found executable at: $exePath" -ForegroundColor Green
Write-Host ""

# Show scheduling options with better formatting
Write-Host "[SCHEDULE] SCHEDULED TASK SETUP:" -ForegroundColor Yellow
Write-Host ""
Write-Host "┌─────────────────────────────────────────────────────────────────┐" -ForegroundColor Gray
Write-Host "│ 1. DAILY ARCHIVING                                              │" -ForegroundColor White
Write-Host "│    Runs once per day at a specific time (e.g., 2:00 AM)         │" -ForegroundColor Gray
Write-Host "│    Best for: Users who want predictable, quiet archiving        │" -ForegroundColor Gray
Write-Host "└─────────────────────────────────────────────────────────────────┘" -ForegroundColor Gray
Write-Host ""

$taskName = "Outlook Auto Archive"

# Daily at specific time
Write-Host ""
Write-Host "Setting up daily scheduled task..." -ForegroundColor Cyan

# Get time from user
Write-Host "What time would you like the script to run daily?" -ForegroundColor Cyan
Write-Host "Recommended: 02:00 (2:00 AM when you're not using Outlook)" -ForegroundColor Gray
do {
    $timeInput = Read-Host "Enter time in 24-hour format (e.g., 02:00 for 2:00 AM)"
    if ($timeInput -match '^([01]?[0-9]|2[0-3]):[0-5][0-9]$') {
        $scheduledTime = $timeInput
        break
    }
    Write-Host "Please enter a valid time in 24-hour format (HH:MM)." -ForegroundColor Red
} while ($true)

# Create daily task
$createCmd = "schtasks /create /tn `"$taskName`" /tr `"$exePath`" /sc daily /st $scheduledTime /f"
Write-Host "Creating daily task..." -ForegroundColor Yellow
Write-Host "Command: $createCmd" -ForegroundColor Gray
Invoke-Expression $createCmd

if ($LASTEXITCODE -eq 0) {
    Write-Host "[OK] Daily scheduled task created successfully!" -ForegroundColor Green
    Write-Host "Task will run daily at $scheduledTime" -ForegroundColor White
}
else {
    Write-Host "[ERROR] Failed to create daily task. Error code: $LASTEXITCODE" -ForegroundColor Red
}

Write-Host ""
Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║   SETUP COMPLETE!                                            ║" -ForegroundColor Green
Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Green
Write-Host ""
Write-Host "[SUCCESS] Your Outlook Auto Archive task has been scheduled successfully!" -ForegroundColor Green
Write-Host ""
Write-Host "[STEPS] NEXT STEPS:" -ForegroundColor Yellow
Write-Host "   • The task will run automatically according to your chosen schedule" -ForegroundColor White
Write-Host "   • Check log files in: C:\Users\$env:USERNAME\OutlookAutoArchive\Logs" -ForegroundColor Cyan
Write-Host ""
Write-Host "[TOOLS] MANAGE YOUR TASK:" -ForegroundColor Yellow
Write-Host "   1. Open Task Scheduler (search in Start menu)" -ForegroundColor White
Write-Host "   2. Look for your task in 'Task Scheduler Library'" -ForegroundColor White
Write-Host "   3. Right-click to modify, disable, or delete the task" -ForegroundColor White
Write-Host ""
Write-Host "[TIP] TIP: The app will gracefully skip runs when Outlook isn't running" -ForegroundColor Cyan
Write-Host ""
Write-Host "Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
