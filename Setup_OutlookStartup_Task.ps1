# Setup Outlook Startup Task
# This script creates a scheduled task that runs the archive script when Outlook starts

param(
    [string]$ScriptPath = $PSScriptRoot,
    [string]$TaskName = "Outlook Auto Archive - On Outlook Start"
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "   Outlook Auto Archive - Startup Task" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check if running as administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "This script requires administrator privileges." -ForegroundColor Red
    Write-Host "Please run PowerShell as Administrator and try again." -ForegroundColor Red
    exit 1
}

# Verify executable exists
$exePath = Join-Path $ScriptPath "OutlookAutoArchive.exe"
if (-not (Test-Path $exePath)) {
    Write-Host "Error: OutlookAutoArchive.exe not found at:" -ForegroundColor Red
    Write-Host $exePath -ForegroundColor Red
    exit 1
}

Write-Host "Found executable at: $exePath" -ForegroundColor Green

# Remove existing task if it exists
Write-Host "Checking for existing task..." -ForegroundColor Yellow
$existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existingTask) {
    Write-Host "Removing existing task: $TaskName" -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

# Create the trigger (when Outlook process starts)
$trigger = New-ScheduledTaskTrigger -AtStartup

# Create the action (run a script that waits for Outlook and then runs the archive)
$actionScript = @"
# Wait for Outlook to start and then run archive
`$maxWait = 300  # 5 minutes
`$waitTime = 0
`$interval = 10  # Check every 10 seconds

Write-Host "Waiting for Outlook to start..."
while (`$waitTime -lt `$maxWait) {
    `$outlookProcesses = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
    if (`$outlookProcesses) {
        Write-Host "Outlook detected. Waiting 30 seconds for it to fully load..."
        Start-Sleep -Seconds 30
        Write-Host "Running archive script..."
        & "$exePath"
        break
    }
    Start-Sleep -Seconds `$interval
    `$waitTime += `$interval
    Write-Host "Still waiting for Outlook... (`$waitTime seconds)"
}

if (`$waitTime -ge `$maxWait) {
    Write-Host "Timeout waiting for Outlook to start."
}
"@

$actionScriptPath = Join-Path $ScriptPath "WaitForOutlook.ps1"
$actionScript | Out-File -FilePath $actionScriptPath -Encoding UTF8

# Create the action
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -File `"$actionScriptPath`"" -WorkingDirectory $ScriptPath

# Create task settings
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

# Create the task
Write-Host "Creating scheduled task..." -ForegroundColor Yellow
Register-ScheduledTask -TaskName $TaskName -Trigger $trigger -Action $action -Settings $settings -Description "Runs Outlook Auto Archive when Outlook starts"

Write-Host ""
Write-Host "Task created successfully!" -ForegroundColor Green
Write-Host "Task Name: $TaskName" -ForegroundColor White
Write-Host "Executable: $exePath" -ForegroundColor White
Write-Host "Trigger: When system starts (waits for Outlook)" -ForegroundColor White
Write-Host ""
Write-Host "How it works:" -ForegroundColor Cyan
Write-Host "1. Task starts when system boots" -ForegroundColor White
Write-Host "2. Waits for Outlook process to start" -ForegroundColor White
Write-Host "3. Waits 30 seconds for Outlook to fully load" -ForegroundColor White
Write-Host "4. Runs the archive script" -ForegroundColor White
Write-Host ""
Write-Host "To test the task:" -ForegroundColor Cyan
Write-Host "1. Open Task Scheduler" -ForegroundColor White
Write-Host "2. Find the task: $TaskName" -ForegroundColor White
Write-Host "3. Right-click and select 'Run'" -ForegroundColor White
Write-Host ""
Write-Host "To modify the task:" -ForegroundColor Cyan
Write-Host "1. Open Task Scheduler" -ForegroundColor White
Write-Host "2. Find the task: $TaskName" -ForegroundColor White
Write-Host "3. Right-click and select 'Properties'" -ForegroundColor White
Write-Host ""
Write-Host "Note: The task will run the archive script each time Outlook starts." -ForegroundColor Yellow
Write-Host "Make sure your config.json has appropriate settings (DryRun, RetentionDays, etc.)." -ForegroundColor Yellow
