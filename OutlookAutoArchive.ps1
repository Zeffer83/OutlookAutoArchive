<#
.SYNOPSIS
  Auto-archive Outlook emails with options from config.json
#>

# Version: 2.8.3
# Author: Ryan Zeffiretti
# Description: Auto-archive Outlook emails with options from config.json

# Try to load Outlook Interop assembly, but don't fail if it's not available
try {
    Add-Type -AssemblyName Microsoft.Office.Interop.Outlook -ErrorAction SilentlyContinue
}
catch {
    Write-Host "Note: Microsoft.Office.Interop.Outlook assembly not found, will use COM objects directly" -ForegroundColor Yellow
}

# Initialize Outlook objects (will be set up later when needed)
$outlook = $null
$namespace = $null

# === Windows Unblocking ===
# Check if this executable was downloaded from the internet and needs to be unblocked
try {
    $currentExePath = if ($PSScriptRoot) { Join-Path $PSScriptRoot "OutlookAutoArchive.exe" } else { Join-Path (Get-Location) "OutlookAutoArchive.exe" }
    
    if (Test-Path $currentExePath) {
        $zoneInfo = Get-ItemProperty -Path $currentExePath -Name Zone.Identifier -ErrorAction SilentlyContinue
        if ($zoneInfo -and $zoneInfo.'Zone.Identifier') {
            Write-Host ""
            Write-Host "⚠️  Windows has blocked this executable because it was downloaded from the internet." -ForegroundColor Yellow
            Write-Host "Attempting to unblock the file automatically..." -ForegroundColor Cyan
            
            try {
                Unblock-File -Path $currentExePath -ErrorAction Stop
                Write-Host "✅ Successfully unblocked the executable!" -ForegroundColor Green
                Write-Host "You can now run the application normally." -ForegroundColor White
            }
            catch {
                Write-Host "❌ Could not automatically unblock the file." -ForegroundColor Red
                Write-Host ""
                Write-Host "To unblock manually:" -ForegroundColor Cyan
                Write-Host "1. Right-click on OutlookAutoArchive.exe" -ForegroundColor White
                Write-Host "2. Select 'Properties'" -ForegroundColor White
                Write-Host "3. Check 'Unblock' at the bottom of the dialog" -ForegroundColor White
                Write-Host "4. Click 'OK'" -ForegroundColor White
                Write-Host "5. Run the application again" -ForegroundColor White
                Write-Host ""
                Write-Host "Press any key to continue anyway..." -ForegroundColor Gray
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
        }
    }
}
catch {
    Write-Host "Note: Could not check for Windows blocking status" -ForegroundColor Gray
}

# === Load config ===
# Handle path for both script and executable
if ($PSScriptRoot) {
    $scriptDir = $PSScriptRoot
}
else {
    # For executable, use current directory
    $scriptDir = Get-Location
}

$configPath = Join-Path $scriptDir 'config.json'
$exampleConfigPath = Join-Path $scriptDir 'config.example.json'

Write-Host "Script directory: $scriptDir"
Write-Host "Config path: $configPath"

# Auto-create config file if missing
if (-not (Test-Path $configPath)) {
    Write-Host "Config file not found. Attempting to create one..."
    
    # Try to copy from example first
    if (Test-Path $exampleConfigPath) {
        Copy-Item $exampleConfigPath $configPath
        Write-Host "Created config.json from config.example.json"
    }
    else {
        # Create default config
        $defaultConfig = @{
            RetentionDays      = 14
            DryRun             = $true
            LogPath            = ".\Logs"
            GmailLabel         = "OutlookArchive"
            OnFirstRun         = $true
            ArchiveFolders     = @{}
            MonitoringInterval = 4  # Hours between continuous monitoring runs
            SkipRules          = @(
                @{
                    Mailbox  = "Your Mailbox Name"
                    Subjects = @("Subject Pattern 1", "Subject Pattern 2")
                }
            )
        }
        
        $defaultConfig | ConvertTo-Json -Depth 3 | Out-File $configPath -Encoding UTF8
        Write-Host "Created default config.json with safe settings (DryRun = true)"
        Write-Host "Please review and edit config.json before running in live mode"
    }
}

try {
    $config = Get-Content $configPath -Raw | ConvertFrom-Json
}
catch {
    Write-Error "Invalid JSON in config.json: $_"
    Write-Host "Please check your config.json file for syntax errors"
    exit 1
}

# === First Run Setup ===
if ($config.OnFirstRun -eq $true) {
    Write-Host ""
    Write-Host "=== Welcome to Outlook Auto Archive - First Run Setup ===" -ForegroundColor Cyan
    Write-Host "This appears to be your first time running the script." -ForegroundColor White
    Write-Host "Let's set up your archive folders and configuration." -ForegroundColor White
    Write-Host ""
    
    # Check admin rights early for scheduling setup
    Write-Host "Checking system requirements..." -ForegroundColor Cyan
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    
    if (-not $isAdmin) {
        Write-Host "⚠️  Note: You're not running as Administrator" -ForegroundColor Yellow
        Write-Host "This is fine for normal usage, but you'll need admin rights for scheduled task creation." -ForegroundColor White
        Write-Host "You can:" -ForegroundColor Cyan
        Write-Host "1. Continue with setup (you can set up scheduling later with admin rights)" -ForegroundColor White
        Write-Host "2. Restart as Administrator now" -ForegroundColor White
        Write-Host ""
        
        do {
            $adminChoice = Read-Host "Continue with setup or restart as Administrator? (1/2)"
            if ($adminChoice -match '^[1-2]$') {
                break
            }
            Write-Host "Please enter 1 or 2." -ForegroundColor Red
        } while ($true)
        
        if ($adminChoice -eq '2') {
            Write-Host "Restarting with admin rights..." -ForegroundColor Yellow
            $scriptPath = if ($PSScriptRoot) { Join-Path $PSScriptRoot "OutlookAutoArchive.exe" } else { Join-Path (Get-Location) "OutlookAutoArchive.exe" }
            Start-Process -FilePath $scriptPath -Verb RunAs
            exit 0
        }
        else {
            Write-Host "Continuing with setup. You can set up scheduling later with admin rights." -ForegroundColor Green
            Write-Host ""
        }
    }
    else {
        Write-Host "✅ Running with Administrator privileges - all features available" -ForegroundColor Green
        Write-Host ""
    }
    
    # Ask about installation location
    Write-Host "Where would you like to install Outlook Auto Archive?" -ForegroundColor Cyan
    Write-Host "This will be the permanent location for the application and its files." -ForegroundColor White
    Write-Host ""
    Write-Host "Recommended locations:" -ForegroundColor Yellow
    Write-Host "1. User Documents (C:\Users\$env:USERNAME\OutlookAutoArchive\) - User-specific installation (Recommended)" -ForegroundColor White
    Write-Host "2. Custom location - Choose your own folder" -ForegroundColor White
    Write-Host "3. Current location - Keep everything where it is now" -ForegroundColor White
    Write-Host ""
    
    do {
        $installChoice = Read-Host "Enter choice (1-3)"
        if ($installChoice -match '^[1-3]$') {
            break
        }
        Write-Host "Please enter 1, 2, or 3." -ForegroundColor Red
    } while ($true)
    
    $installPath = ""
    $currentLocation = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }
    
    if ($installChoice -eq '1') {
        $installPath = "$env:USERPROFILE\OutlookAutoArchive"
        Write-Host "Selected: User Documents installation (Recommended)" -ForegroundColor Green
    }
    elseif ($installChoice -eq '2') {
        Write-Host ""
        Write-Host "Enter the full path where you want to install the application:" -ForegroundColor Cyan
        Write-Host "Example: C:\MyTools\OutlookAutoArchive" -ForegroundColor Gray
        do {
            $customPath = Read-Host "Installation path"
            if ([string]::IsNullOrWhiteSpace($customPath)) {
                Write-Host "Please enter a valid path." -ForegroundColor Red
                continue
            }
            
            # Validate the path
            try {
                $installPath = [System.IO.Path]::GetFullPath($customPath)
                break
            }
            catch {
                Write-Host "Invalid path format. Please enter a valid path." -ForegroundColor Red
            }
        } while ($true)
        Write-Host "Selected: Custom location ($installPath)" -ForegroundColor Green
    }
    elseif ($installChoice -eq '3') {
        $installPath = $currentLocation
        Write-Host "Selected: Current location ($installPath)" -ForegroundColor Green
    }
    
    # Check if we need to move files
    if ($installPath -ne $currentLocation) {
        Write-Host ""
        Write-Host "Setting up installation at: $installPath" -ForegroundColor Cyan
        
        try {
            # Create the installation directory if it doesn't exist
            if (-not (Test-Path $installPath)) {
                New-Item -Path $installPath -ItemType Directory -Force | Out-Null
                Write-Host "✅ Created installation directory" -ForegroundColor Green
            }
            
            # Copy only essential files to the installation location
            $filesToCopy = @(
                "OutlookAutoArchive.exe",
                "config.example.json"
            )
            
            $filesCopied = 0
            foreach ($file in $filesToCopy) {
                $sourceFile = Join-Path $currentLocation $file
                $destFile = Join-Path $installPath $file
                
                if (Test-Path $sourceFile) {
                    Copy-Item -Path $sourceFile -Destination $destFile -Force
                    $filesCopied++
                }
            }
            
            Write-Host "✅ Copied $filesCopied files to installation directory" -ForegroundColor Green
            
            # Create a simple README.txt for users
            $readmeContent = @"
Outlook Auto Archive - Version 2.2.0
====================================

This application automatically archives old emails from your Outlook accounts.

QUICK START:
1. Double-click OutlookAutoArchive.exe to run
2. If Windows blocks the file, the app will attempt to unblock it automatically
3. Follow the setup wizard to configure your preferences
4. The app will create a config.json file with your settings
5. Check the Logs folder for operation details

WINDOWS SECURITY:
- If Windows blocks the executable, the app will try to unblock it automatically
- If automatic unblocking fails, right-click the .exe file → Properties → Check "Unblock"
- This is normal for files downloaded from the internet

CONFIGURATION:
- Edit config.json to change settings (retention days, dry-run mode, etc.)
- Set 'DryRun': false when ready to archive emails for real
- Logs are stored in the Logs folder within this directory

SUPPORT:
- For help and updates, visit the original repository
- Check the Logs folder for troubleshooting information

Version 2.2.0 - Professional metadata and Windows security handling
"@
            
            $readmePath = Join-Path $installPath "README.txt"
            $readmeContent | Out-File -FilePath $readmePath -Encoding UTF8
            Write-Host "✅ Created user-friendly README.txt" -ForegroundColor Green
            
            # Update the script directory for the rest of the setup
            $scriptDir = $installPath
            $configPath = Join-Path $installPath 'config.json'
            $exampleConfigPath = Join-Path $installPath 'config.example.json'
            
            Write-Host ""
            Write-Host "Installation completed successfully!" -ForegroundColor Green
            Write-Host "The application is now installed at: $installPath" -ForegroundColor White
            Write-Host ""
            Write-Host "Note: You can now delete the original files from: $currentLocation" -ForegroundColor Yellow
            Write-Host "The application will run from the new location." -ForegroundColor White
            Write-Host ""
            
        }
        catch {
            Write-Host "❌ Error during installation: $_" -ForegroundColor Red
            Write-Host "Continuing with current location..." -ForegroundColor Yellow
            $scriptDir = $currentLocation
            $configPath = Join-Path $currentLocation 'config.json'
            $exampleConfigPath = Join-Path $currentLocation 'config.example.json'
        }
    }
    else {
        # Keep current location
        $scriptDir = $currentLocation
        $configPath = Join-Path $currentLocation 'config.json'
        $exampleConfigPath = Join-Path $currentLocation 'config.example.json'
    }
    
    # Check if Outlook is running
    try {
        $outlookProcesses = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
        if (-not $outlookProcesses) {
            Write-Host "❌ Outlook is not running. Please start Outlook and run the script again." -ForegroundColor Red
            Write-Host "The setup requires Outlook to be running to access your email accounts." -ForegroundColor Yellow
            exit 1
        }
        Write-Host "✅ Outlook is running" -ForegroundColor Green
    }
    catch {
        Write-Host "⚠️  Could not check Outlook status. Proceeding anyway..." -ForegroundColor Yellow
    }
    
    # Connect to Outlook
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        Write-Host "✅ Connected to Outlook" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Failed to connect to Outlook: $_" -ForegroundColor Red
        Write-Host "Make sure Outlook is running and you have the necessary permissions." -ForegroundColor Yellow
        exit 1
    }
    
    Write-Host ""
    Write-Host "Scanning your email accounts..." -ForegroundColor Cyan
    
    $accounts = @()
    $gmailAccounts = @()
    $regularAccounts = @()
    
    foreach ($account in $namespace.Folders) {
        $accounts += $account.Name
        
        # Check if this looks like a Gmail account
        $isGmail = $account.Name -like "*@gmail.com" -or $account.Name -like "*@googlemail.com" -or $account.Name -like "*@gmail.co.uk"
        
        if ($isGmail) {
            $gmailAccounts += $account.Name
        }
        else {
            $regularAccounts += $account.Name
        }
    }
    
    Write-Host "Found $($accounts.Count) email account(s):" -ForegroundColor Green
    foreach ($account in $accounts) {
        Write-Host "  - $account" -ForegroundColor White
    }
    
    if ($gmailAccounts.Count -gt 0) {
        Write-Host ""
        Write-Host "Gmail accounts detected: $($gmailAccounts -join ', ')" -ForegroundColor Yellow
        Write-Host "Note: Gmail accounts will use labels instead of folders for archiving." -ForegroundColor Gray
    }
    
    Write-Host ""
    
    # Configure retention days
    Write-Host "How many days should emails stay in your Inbox before being archived?" -ForegroundColor Cyan
    Write-Host "Recommended: 14-30 days" -ForegroundColor Gray
    do {
        $retentionInput = Read-Host "Enter number of days (default: 14)"
        if ([string]::IsNullOrWhiteSpace($retentionInput)) {
            $retentionDays = 14
            break
        }
        if ($retentionInput -match '^\d+$' -and [int]$retentionInput -gt 0) {
            $retentionDays = [int]$retentionInput
            break
        }
        Write-Host "Please enter a valid positive number." -ForegroundColor Red
    } while ($true)
    
    Write-Host "✅ Retention period set to $retentionDays days" -ForegroundColor Green
    
    # Configure Gmail label if Gmail accounts exist
    $gmailLabel = "OutlookArchive"
    if ($gmailAccounts.Count -gt 0) {
        Write-Host ""
        Write-Host "For Gmail accounts, what would you like to call your archive label?" -ForegroundColor Cyan
        Write-Host "Note: 'Archive' is not allowed in Gmail, so we use a custom label name." -ForegroundColor Gray
        Write-Host "Recommended: OutlookArchive, MyArchive, or EmailArchive" -ForegroundColor Gray
        do {
            $labelInput = Read-Host "Enter label name (default: OutlookArchive)"
            if ([string]::IsNullOrWhiteSpace($labelInput)) {
                $gmailLabel = "OutlookArchive"
                break
            }
            if ($labelInput -match '^[a-zA-Z0-9_-]+$') {
                $gmailLabel = $labelInput
                break
            }
            Write-Host "Please enter a valid label name (letters, numbers, hyphens, underscores only)." -ForegroundColor Red
        } while ($true)
        
        Write-Host "✅ Gmail archive label set to '$gmailLabel'" -ForegroundColor Green
    }
    
    Write-Host ""
    Write-Host "Now let's check for existing archive folders and create any missing ones..." -ForegroundColor Cyan
    
    $foldersCreated = 0
    $errors = 0
    
    foreach ($account in $namespace.Folders) {
        try {
            Write-Host ""
            Write-Host "Processing account: $($account.Name)" -ForegroundColor Cyan
            
            # Skip non-email account types
            $skipAccountTypes = @("Internet Calendars", "SharePoint Lists", "Public Folders", "Calendar", "Contacts", "Tasks", "Notes")
            if ($skipAccountTypes -contains $account.Name) {
                Write-Host "  ⚠️  Skipping non-email account type: $($account.Name)" -ForegroundColor Yellow
                continue
            }
            
            # Check if this looks like a Gmail account
            $isGmail = $account.Name -like "*@gmail.com" -or $account.Name -like "*@googlemail.com" -or $account.Name -like "*@gmail.co.uk"
            
            if ($isGmail) {
                Write-Host "  Detected Gmail account" -ForegroundColor Gray
                
                # Check if Gmail label already exists
                $existingLabel = $null
                try {
                    $existingLabel = $account.Folders.Item($gmailLabel)
                }
                catch {}
                
                if ($existingLabel) {
                    Write-Host "  ✅ Gmail label '$gmailLabel' already exists" -ForegroundColor Green
                    # Store the Gmail label path in config
                    $config.ArchiveFolders[$account.Name] = "GmailLabel:$gmailLabel"
                }
                else {
                    Write-Host "  Gmail label '$gmailLabel' not found" -ForegroundColor Yellow
                    $createLabel = Read-Host "  Would you like to create it? (Y/N)"
                    if ($createLabel -eq 'Y' -or $createLabel -eq 'y') {
                        try {
                            $account.Folders.Add($gmailLabel)
                            Write-Host "  ✅ Created Gmail label '$gmailLabel'" -ForegroundColor Green
                            $foldersCreated++
                            # Store the Gmail label path in config
                            $config.ArchiveFolders[$account.Name] = "GmailLabel:$gmailLabel"
                        }
                        catch {
                            # Check if the label was actually created despite the error
                            try {
                                $testLabel = $account.Folders.Item($gmailLabel)
                                if ($testLabel) {
                                    Write-Host "  ✅ Gmail label '$gmailLabel' was created successfully" -ForegroundColor Green
                                    $foldersCreated++
                                    # Store the Gmail label path in config
                                    $config.ArchiveFolders[$account.Name] = "GmailLabel:$gmailLabel"
                                }
                            }
                            catch {
                                Write-Host "  ⚠️  Gmail label creation encountered an issue, but this is often normal for Gmail accounts" -ForegroundColor Yellow
                                Write-Host "  The label may still be available in Outlook. You can check manually or try again later." -ForegroundColor Gray
                                $errors++
                            }
                        }
                    }
                    else {
                        Write-Host "  ⚠️  Skipped creating Gmail label" -ForegroundColor Yellow
                    }
                }
            }
            else {
                Write-Host "  Detected regular email account" -ForegroundColor Gray
                
                # Check for existing archive folder
                $archiveFolder = $null
                
                # Check root level first
                try {
                    $archiveFolder = $account.Folders.Item("Archive")
                    Write-Host "  ✅ Archive folder already exists at root level" -ForegroundColor Green
                    # Store the archive folder path in config
                    $config.ArchiveFolders[$account.Name] = "Root:Archive"
                }
                catch {
                    # Check Inbox\Archive
                    try {
                        $inbox = $account.Folders.Item("Inbox")
                        if ($inbox) {
                            $archiveFolder = $inbox.Folders.Item("Archive")
                            Write-Host "  ✅ Archive folder already exists in Inbox" -ForegroundColor Green
                            # Store the archive folder path in config
                            $config.ArchiveFolders[$account.Name] = "Inbox:Archive"
                        }
                    }
                    catch {}
                    
                    # If no archive folder found, ask user where to create it
                    if (-not $archiveFolder) {
                        Write-Host "  No Archive folder found" -ForegroundColor Yellow
                        Write-Host "  Where would you like to create the Archive folder?" -ForegroundColor Cyan
                        Write-Host "  1. Root level (recommended)" -ForegroundColor White
                        Write-Host "  2. Inside Inbox folder" -ForegroundColor White
                        Write-Host "  3. Skip this account" -ForegroundColor White
                        
                        do {
                            $locationChoice = Read-Host "  Enter choice (1-3)"
                            if ($locationChoice -match '^[1-3]$') {
                                break
                            }
                            Write-Host "  Please enter 1, 2, or 3." -ForegroundColor Red
                        } while ($true)
                        
                        if ($locationChoice -eq '1') {
                            try {
                                $archiveFolder = $account.Folders.Add("Archive")
                                Write-Host "  ✅ Created Archive folder at root level" -ForegroundColor Green
                                $foldersCreated++
                                # Store the archive folder path in config
                                $config.ArchiveFolders[$account.Name] = "Root:Archive"
                            }
                            catch {
                                Write-Host "  ❌ Failed to create Archive folder: $_" -ForegroundColor Red
                                $errors++
                            }
                        }
                        elseif ($locationChoice -eq '2') {
                            try {
                                $inbox = $account.Folders.Item("Inbox")
                                if ($inbox) {
                                    $archiveFolder = $inbox.Folders.Add("Archive")
                                    Write-Host "  ✅ Created Archive folder in Inbox" -ForegroundColor Green
                                    $foldersCreated++
                                    # Store the archive folder path in config
                                    $config.ArchiveFolders[$account.Name] = "Inbox:Archive"
                                }
                                else {
                                    Write-Host "  ❌ Could not access Inbox folder" -ForegroundColor Red
                                    $errors++
                                }
                            }
                            catch {
                                Write-Host "  ❌ Failed to create Archive folder: $_" -ForegroundColor Red
                                $errors++
                            }
                        }
                        else {
                            Write-Host "  ⚠️  Skipped creating Archive folder" -ForegroundColor Yellow
                        }
                    }
                }
            }
        }
        catch {
            Write-Host "  ❌ Error processing account '$($account.Name)': $_" -ForegroundColor Red
            $errors++
        }
    }
    
    Write-Host ""
    Write-Host "=== Setup Summary ===" -ForegroundColor Cyan
    Write-Host "Accounts processed: $($accounts.Count)" -ForegroundColor White
    Write-Host "Folders/labels created: $foldersCreated" -ForegroundColor White
    Write-Host "Errors encountered: $errors" -ForegroundColor $(if ($errors -gt 0) { "Red" } else { "Green" })
    
    # Update config with user preferences
    $config.RetentionDays = $retentionDays
    $config.GmailLabel = $gmailLabel
    $config.OnFirstRun = $false
     
    # Save monitoring interval if it was set
    if ($monitoringInterval) {
        $config.MonitoringInterval = $monitoringInterval
    }
    
    # Save updated config with discovered archive folders
    try {
        $config | ConvertTo-Json -Depth 3 | Out-File $configPath -Encoding UTF8
        Write-Host "✅ Configuration saved with archive folder paths" -ForegroundColor Green
        Write-Host "Archive folders discovered and stored for future runs:" -ForegroundColor Cyan
        foreach ($accountName in $config.ArchiveFolders.Keys) {
            Write-Host "  - $accountName`: $($config.ArchiveFolders[$accountName])" -ForegroundColor White
        }
    }
    catch {
        Write-Host "❌ Failed to save configuration: $_" -ForegroundColor Red
    }
    
    Write-Host ""
    Write-Host "🎉 First run setup completed!" -ForegroundColor Green
    Write-Host ""
     
    # Ask about scheduled task setup
    Write-Host "Would you like to set up automatic scheduling now?" -ForegroundColor Cyan
    Write-Host "This will create a Windows Task Scheduler task to run the archive script automatically." -ForegroundColor White
    Write-Host ""
    Write-Host "Scheduling options:" -ForegroundColor Yellow
    Write-Host "1. Daily at a specific time (e.g., 2:00 AM)" -ForegroundColor White
    Write-Host "2. When Outlook starts + every 4 hours (recommended)" -ForegroundColor White
    Write-Host "3. Skip scheduling for now" -ForegroundColor White
    Write-Host ""
     
    do {
        $scheduleChoice = Read-Host "Enter choice (1-3)"
        if ($scheduleChoice -match '^[1-3]$') {
            break
        }
        Write-Host "Please enter 1, 2, or 3." -ForegroundColor Red
    } while ($true)
     
    if ($scheduleChoice -eq '1') {
        Write-Host ""
        Write-Host "Setting up daily scheduled task..." -ForegroundColor Cyan
         
        # Get time from user
        Write-Host "What time would you like the script to run daily?" -ForegroundColor Cyan
        Write-Host "Recommended: 2:00 AM (when you're not using Outlook)" -ForegroundColor Gray
        do {
            $timeInput = Read-Host "Enter time in 24-hour format (e.g., 02:00 for 2:00 AM)"
            if ($timeInput -match '^([01]?[0-9]|2[0-3]):[0-5][0-9]$') {
                $scheduledTime = $timeInput
                break
            }
            Write-Host "Please enter a valid time in 24-hour format (HH:MM)." -ForegroundColor Red
        } while ($true)
         
        # Create daily scheduled task (admin rights already checked at setup start)
        if (-not $isAdmin) {
            Write-Host "⚠️  Admin rights required for scheduled task creation" -ForegroundColor Yellow
            Write-Host "You can set up scheduling manually later using Task Scheduler:" -ForegroundColor White
            Write-Host "1. Open Task Scheduler (search in Start menu)" -ForegroundColor Gray
            Write-Host "2. Click 'Create Basic Task'" -ForegroundColor Gray
            Write-Host "3. Name: 'Outlook Auto Archive'" -ForegroundColor Gray
            Write-Host "4. Trigger: 'Daily' at $scheduledTime" -ForegroundColor Gray
            Write-Host "5. Action: 'Start a program'" -ForegroundColor Gray
            Write-Host "6. Program: '$scriptPath'" -ForegroundColor Gray
            Write-Host "7. Finish and check 'Open properties dialog'" -ForegroundColor Gray
            Write-Host "8. In Properties, go to 'General' tab and check 'Run with highest privileges'" -ForegroundColor Gray
            Write-Host "9. Click OK to save" -ForegroundColor Gray
            Write-Host ""
            Write-Host "The task will run daily at $scheduledTime." -ForegroundColor Green
        }
        else {
            # Create daily scheduled task
            try {
                $taskName = "Outlook Auto Archive"
                $scriptPath = if ($PSScriptRoot) { Join-Path $PSScriptRoot "OutlookAutoArchive.exe" } else { Join-Path (Get-Location) "OutlookAutoArchive.exe" }
                 
                # Check if executable exists, fall back to PowerShell script
                if (-not (Test-Path $scriptPath)) {
                    $scriptPath = if ($PSScriptRoot) { Join-Path $PSScriptRoot "OutlookAutoArchive.ps1" } else { Join-Path (Get-Location) "OutlookAutoArchive.ps1" }
                    $arguments = "-ExecutionPolicy Bypass -File `"$scriptPath`""
                    $program = "powershell.exe"
                }
                else {
                    $arguments = ""
                    $program = $scriptPath
                }
                 
                # Create the scheduled task
                $createTaskCmd = "schtasks /create /tn `"$taskName`" /tr `"$program`""
                if ($arguments) { $createTaskCmd += " /sc daily /st $scheduledTime /f" } else { $createTaskCmd += " /sc daily /st $scheduledTime /f" }
                 
                Write-Host "Creating scheduled task..." -ForegroundColor Yellow
                Invoke-Expression $createTaskCmd
                 
                if ($LASTEXITCODE -eq 0) {
                    Write-Host "✅ Daily scheduled task created successfully!" -ForegroundColor Green
                    Write-Host "Task will run daily at $scheduledTime" -ForegroundColor White
                }
                else {
                    Write-Host "⚠️  Could not create scheduled task automatically." -ForegroundColor Yellow
                    Write-Host "You can create it manually using Task Scheduler:" -ForegroundColor White
                    Write-Host "1. Open Task Scheduler" -ForegroundColor Gray
                    Write-Host "2. Create Basic Task" -ForegroundColor Gray
                    Write-Host "3. Name: Outlook Auto Archive" -ForegroundColor Gray
                    Write-Host "4. Trigger: Daily at $scheduledTime" -ForegroundColor Gray
                    Write-Host "5. Action: Start program: $program" -ForegroundColor Gray
                }
            }
            catch {
                Write-Host "❌ Error creating scheduled task: $_" -ForegroundColor Red
                Write-Host "You can set up scheduling manually later." -ForegroundColor Yellow
            }
        }
    }
    elseif ($scheduleChoice -eq '2') {
        Write-Host ""
        Write-Host "Setting up Outlook startup + periodic monitoring..." -ForegroundColor Cyan
        Write-Host "This creates a task that starts when Outlook opens, then runs every 4 hours." -ForegroundColor White
        Write-Host "Perfect for users who want archiving only when Outlook is available!" -ForegroundColor Gray
        Write-Host ""
         
        # Create startup + monitoring task (admin rights already checked at setup start)
        if (-not $isAdmin) {
            Write-Host "⚠️  Admin rights required for scheduled task creation" -ForegroundColor Yellow
            Write-Host "You can set up scheduling manually later using Task Scheduler:" -ForegroundColor White
            Write-Host "1. Open Task Scheduler (search in Start menu)" -ForegroundColor Gray
            Write-Host "2. Click 'Create Basic Task'" -ForegroundColor Gray
            Write-Host "3. Name: 'Outlook Auto Archive - Startup + Monitoring'" -ForegroundColor Gray
            Write-Host "4. Trigger: 'When the computer starts'" -ForegroundColor Gray
            Write-Host "5. Action: 'Start a program'" -ForegroundColor Gray
            Write-Host "6. Program: '$scriptPath'" -ForegroundColor Gray
            Write-Host "7. Finish and check 'Open properties dialog'" -ForegroundColor Gray
            Write-Host "8. In Properties, go to 'Triggers' tab and edit the trigger" -ForegroundColor Gray
            Write-Host "9. Set 'Repeat task every: 4 hours'" -ForegroundColor Gray
            Write-Host "10. Set 'for a duration of: Indefinitely'" -ForegroundColor Gray
            Write-Host "11. In 'General' tab, check 'Run with highest privileges'" -ForegroundColor Gray
            Write-Host "12. Click OK to save" -ForegroundColor Gray
            Write-Host ""
            Write-Host "The task will start when the computer starts and run every 4 hours." -ForegroundColor Green
            Write-Host "The script will gracefully skip runs when Outlook is not available." -ForegroundColor Green
            Write-Host ""
        }
        else {
            try {
                $taskName = "Outlook Auto Archive - Startup + Monitoring"
                $scriptPath = if ($PSScriptRoot) { Join-Path $PSScriptRoot "OutlookAutoArchive.exe" } else { Join-Path (Get-Location) "OutlookAutoArchive.exe" }
                 
                if (-not (Test-Path $scriptPath)) {
                    $scriptPath = if ($PSScriptRoot) { Join-Path $PSScriptRoot "OutlookAutoArchive.ps1" } else { Join-Path (Get-Location) "OutlookAutoArchive.ps1" }
                    $arguments = "-ExecutionPolicy Bypass -File `"$scriptPath`""
                    $program = "powershell.exe"
                }
                else {
                    $arguments = ""
                    $program = $scriptPath
                }
                 
                # Create the startup + monitoring task
                # This will start when the computer starts and run every 4 hours
                $createTaskCmd = "schtasks /create /tn `"$taskName`" /tr `"$program`" /sc onstart /delay 0000:30 /mo 4 /f"
                if ($arguments) { 
                    $createTaskCmd = "schtasks /create /tn `"$taskName`" /tr `"$program $arguments`" /sc onstart /delay 0000:30 /mo 4 /f"
                }
                 
                Write-Host "Creating startup + monitoring task: $taskName" -ForegroundColor Yellow
                Invoke-Expression $createTaskCmd
                 
                if ($LASTEXITCODE -eq 0) {
                    Write-Host "✅ Startup + monitoring task created successfully!" -ForegroundColor Green
                    Write-Host "Task name: $taskName" -ForegroundColor White
                    Write-Host "Task will start 30 seconds after system startup and run every 4 hours" -ForegroundColor White
                    Write-Host "The script will gracefully skip runs when Outlook is not available" -ForegroundColor White
                    Write-Host "You can find it in Task Scheduler under 'Task Scheduler Library'" -ForegroundColor Cyan
                    Write-Host ""
                    Write-Host "This is the best option for users who want archiving only when Outlook is available!" -ForegroundColor Green
                }
                else {
                    Write-Host "⚠️  Could not create startup + monitoring task automatically." -ForegroundColor Yellow
                    Write-Host "Error code: $LASTEXITCODE" -ForegroundColor Red
                    Write-Host "You can create it manually in Task Scheduler:" -ForegroundColor White
                    Write-Host "1. Open Task Scheduler" -ForegroundColor Gray
                    Write-Host "2. Create Basic Task" -ForegroundColor Gray
                    Write-Host "3. Name: $taskName" -ForegroundColor Gray
                    Write-Host "4. Trigger: At system startup" -ForegroundColor Gray
                    Write-Host "5. Action: Start program: $program" -ForegroundColor Gray
                    if ($arguments) {
                        Write-Host "6. Arguments: $arguments" -ForegroundColor Gray
                    }
                    Write-Host "7. In Properties, edit trigger to repeat every 4 hours" -ForegroundColor Gray
                }
            }
            catch {
                Write-Host "❌ Error creating startup + monitoring task: $_" -ForegroundColor Red
                Write-Host "You can set up scheduling manually later." -ForegroundColor Yellow
            }
        }
    }
    else {
        Write-Host ""
        Write-Host "Scheduling skipped. You can set it up later using:" -ForegroundColor Yellow
        Write-Host "1. Task Scheduler GUI" -ForegroundColor White
        Write-Host "2. Setup_OutlookStartup_Task.ps1 script" -ForegroundColor White
        Write-Host "3. Manual schtasks command" -ForegroundColor White
    }
     
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Cyan
    Write-Host "1. The script will now run in dry-run mode to test everything" -ForegroundColor White
    Write-Host "2. Check the log files to verify everything works" -ForegroundColor White
    Write-Host "3. When ready, edit config.json and set 'DryRun': false" -ForegroundColor White
    Write-Host "4. Test your scheduled task if you created one" -ForegroundColor White
     
    Write-Host ""
    Write-Host "⚠️  IMPORTANT: The dry-run test may take several minutes depending on how many emails you have." -ForegroundColor Yellow
    Write-Host "This is normal - the script is scanning all your emails to show what would be archived." -ForegroundColor White
    Write-Host "Please be patient and don't close the window while it's running." -ForegroundColor White
    
    if ($gmailAccounts.Count -gt 0) {
        Write-Host ""
        Write-Host "For Gmail users:" -ForegroundColor Cyan
        Write-Host "- Make sure IMAP is enabled in Gmail settings" -ForegroundColor White
        Write-Host "- Check 'Show in IMAP' for your labels in Gmail web interface" -ForegroundColor White
        Write-Host "- It may take a few minutes for labels to sync to Outlook" -ForegroundColor White
    }
    
    Write-Host ""
    Write-Host "=== Continuing with archive process... ===" -ForegroundColor Cyan
    Write-Host ""
    
    # Reload config to get the updated values from the first-run setup
    try {
        $config = Get-Content $configPath -Raw | ConvertFrom-Json
        Write-Host "✅ Configuration reloaded with updated settings" -ForegroundColor Green
    }
    catch {
        Write-Host "⚠️  Could not reload configuration, continuing with current settings" -ForegroundColor Yellow
    }
}

# === Apply config settings ===
$RetentionDays = [int]$config.RetentionDays
$DryRun = [bool]$config.DryRun

# Process log path with proper error handling
$rawLogPath = $config.LogPath
if ([string]::IsNullOrEmpty($rawLogPath)) {
    $rawLogPath = ".\Logs"
    Write-Host "LogPath was empty, using default: $rawLogPath"
}

# Handle relative paths and environment variables
if ($rawLogPath -like ".\*") {
    # Relative path - make it absolute based on script location
    $LogPath = Join-Path $scriptDir $rawLogPath.Substring(2)
}
else {
    # Handle both escaped and unescaped backslashes for absolute paths
    $LogPath = $rawLogPath -replace '%USERPROFILE%', $env:USERPROFILE
    $LogPath = $LogPath -replace '\\\\', '\'  # Fix double backslashes
}

if ([string]::IsNullOrEmpty($LogPath)) {
    $LogPath = Join-Path $scriptDir "Logs"
    Write-Host "LogPath processing failed, using fallback: $LogPath"
}

Write-Host "Using log path: $LogPath"

$Today = Get-Date
$CutOff = $Today.AddDays(-$RetentionDays)
$GmailLabel = $config.GmailLabel
$SkipRules = $config.SkipRules

# === Check if Outlook is running (for interactive runs only) ===
# Note: Scheduled runs handle Outlook availability gracefully in the connection section below
if ([Environment]::UserInteractive) {
    try {
        $outlookProcesses = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
        if (-not $outlookProcesses) {
            Write-Host "❌ Outlook is not running!" -ForegroundColor Red
            Write-Host ""
            Write-Host "The archive script requires Outlook to be running to access email data." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "Please:" -ForegroundColor Cyan
            Write-Host "1. Start Outlook manually" -ForegroundColor White
            Write-Host "2. Run this script again" -ForegroundColor White
            Write-Host "3. Set up a scheduled task that runs when Outlook starts" -ForegroundColor White
            Write-Host ""
            Write-Host "Press any key to exit..." -ForegroundColor Gray
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            exit 1
        }
        Write-Host "✅ Outlook is running. Proceeding with archive process..." -ForegroundColor Green
    }
    catch {
        Write-Host "⚠️  Could not check Outlook status. Proceeding anyway..." -ForegroundColor Yellow
    }
}

# === Setup logging ===
$LogFile = $null
try {
    # Ensure LogPath is valid
    if ([string]::IsNullOrEmpty($LogPath)) {
        throw "LogPath is null or empty"
    }
    
    # Test if we can create the directory
    if (-not (Test-Path $LogPath)) { 
        New-Item -Path $LogPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
        Write-Host "Created log directory: $LogPath"
    }
    
    # Create log file path
    $LogFile = Join-Path $LogPath ("ArchiveLog_" + $Today.ToString("yyyy-MM-dd_HH-mm-ss") + ".txt")
    
    # Test writing to the log file
    "=== Outlook Auto-Archive Dry-Run ===" | Out-File -FilePath $LogFile -Encoding UTF8
    "Retention: $RetentionDays days"       | Out-File -FilePath $LogFile -Append -Encoding UTF8
    "Cutoff: $CutOff"                       | Out-File -FilePath $LogFile -Append -Encoding UTF8
    
    Write-Host "Logging initialized successfully: $LogFile"
}
catch {
    Write-Host "Error setting up logging: $_" -ForegroundColor Red
    Write-Host "LogPath: $LogPath" -ForegroundColor Yellow
    Write-Host "Continuing without logging..." -ForegroundColor Yellow
    $LogFile = $null
}

# === Connect to Outlook for main processing ===
if (-not $outlook -or -not $namespace) {
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        Write-Host "✅ Connected to Outlook for processing" -ForegroundColor Green
    }
    catch {
        # Check if this is a scheduled run (non-interactive)
        $isScheduledRun = $false
        try {
            # Check if we're running from Task Scheduler (non-interactive environment)
            $isScheduledRun = -not [Environment]::UserInteractive -or $null -eq $env:COMPUTERNAME
        }
        catch {
            # If we can't determine, assume it might be scheduled
            $isScheduledRun = $true
        }
        
        if ($isScheduledRun) {
            # Graceful handling for scheduled runs
            $logMessage = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Outlook is not running. Skipping scheduled archive run."
            Write-Host $logMessage -ForegroundColor Yellow
            
            # Try to log to file if possible
            if ($LogFile -and (Test-Path (Split-Path $LogFile -Parent))) {
                try {
                    $logMessage | Out-File -FilePath $LogFile -Append -Encoding UTF8 -ErrorAction SilentlyContinue
                }
                catch {
                    # Silently continue if logging fails
                }
            }
            
            # Exit gracefully with success code for scheduled tasks
            exit 0
        }
        else {
            # Interactive run - show error and exit with failure
            Write-Host "❌ Failed to connect to Outlook: $_" -ForegroundColor Red
            Write-Host "Make sure Outlook is running and you have the necessary permissions." -ForegroundColor Yellow
            exit 1
        }
    }
}

# Helper function for safe logging
function Write-Log {
    param(
        [string]$Message,
        [string]$LogFile
    )
    
    # Always write to console
    Write-Host $Message
    
    # Only write to file if LogFile is not null and exists
    if ($LogFile -and (Test-Path $LogFile)) {
        try {
            $Message | Out-File -FilePath $LogFile -Append -Encoding UTF8 -ErrorAction SilentlyContinue
        }
        catch {
            Write-Host "Warning: Could not write to log file: $_" -ForegroundColor Yellow
        }
    }
}

function Get-ArchiveFolder {
    param($account)

    $archive = $null

    # Check if we have a stored path for this account
    if ($config.ArchiveFolders -and (Get-Member -InputObject $config.ArchiveFolders -Name $account.Name)) {
        $storedPath = $config.ArchiveFolders[$account.Name]
        
        if ($storedPath -like "GmailLabel:*") {
            # Gmail label path
            $labelName = $storedPath.Split(":")[1]
            try {
                $archive = $account.Folders.Item($labelName)
                Write-Host "  Using stored Gmail label: $labelName" -ForegroundColor Green
            }
            catch {
                Write-Host "  Stored Gmail label '$labelName' not found, will search..." -ForegroundColor Yellow
            }
        }
        elseif ($storedPath -like "Root:*") {
            # Root-level archive folder
            $folderName = $storedPath.Split(":")[1]
            try {
                $archive = $account.Folders.Item($folderName)
                Write-Host "  Using stored root folder: $folderName" -ForegroundColor Green
            }
            catch {
                Write-Host "  Stored root folder '$folderName' not found, will search..." -ForegroundColor Yellow
            }
        }
        elseif ($storedPath -like "Inbox:*") {
            # Inbox-level archive folder
            $folderName = $storedPath.Split(":")[1]
            try {
                $inbox = $account.Folders.Item("Inbox")
                if ($inbox) {
                    $archive = $inbox.Folders.Item($folderName)
                    Write-Host "  Using stored Inbox folder: Inbox\$folderName" -ForegroundColor Green
                }
            }
            catch {
                Write-Host "  Stored Inbox folder 'Inbox\$folderName' not found, will search..." -ForegroundColor Yellow
            }
        }
    }

    # If stored path didn't work, fall back to search (for backward compatibility)
    if (-not $archive) {
        Write-Host "  Searching for archive folders..." -ForegroundColor Yellow
        
        # Inbox\Archive
        try {
            $inbox = $account.Folders.Item("Inbox")
            if ($inbox -and ($inbox.Folders | Where-Object { $_.Name -eq "Archive" })) {
                $archive = $inbox.Folders.Item("Archive")
                Write-Host "  Found Inbox\Archive folder" -ForegroundColor Green
            }
        }
        catch {}

        # Root-level Archive
        if (-not $archive) {
            if ($account.Folders | Where-Object { $_.Name -eq "Archive" }) {
                try { 
                    $archive = $account.Folders.Item("Archive") 
                    Write-Host "  Found root Archive folder" -ForegroundColor Green
                }
                catch {}
            }
        }

        # Gmail custom label
        if (-not $archive -and $GmailLabel) {
            Write-Host "  Looking for Gmail label: $GmailLabel" -ForegroundColor Yellow
            try {
                # Try to access the Gmail label directly by name
                try {
                    $archive = $account.Folders.Item($GmailLabel)
                    Write-Host "  Found Gmail label: $GmailLabel" -ForegroundColor Green
                }
                catch {
                    Write-Host "  Gmail label '$GmailLabel' not found" -ForegroundColor Yellow
                    
                    # Try to enumerate folders as fallback
                    $folders = @($account.Folders)
                    Write-Host "  Found $($folders.Count) folders" -ForegroundColor Gray
                    
                    foreach ($folder in $folders) {
                        Write-Host "  - $($folder.Name)" -ForegroundColor Gray
                        if ($folder.Name -eq $GmailLabel) { 
                            $archive = $folder; 
                            Write-Host "  Found Gmail label: $GmailLabel" -ForegroundColor Green
                            break 
                        }
                    }
                }
            }
            catch {
                Write-Host "  Error accessing folders: $_" -ForegroundColor Red
            }
        }
    }

    return $archive
}

foreach ($account in $namespace.Folders) {
    try {
        Write-Host "Processing account: $($account.Name)" -ForegroundColor Cyan
        
        # Skip non-email account types
        $skipAccountTypes = @("Internet Calendars", "SharePoint Lists", "Public Folders", "Calendar", "Contacts", "Tasks", "Notes")
        if ($skipAccountTypes -contains $account.Name) {
            $logMessage = "[$($account.Name)] Skipping non-email account type."
            Write-Log -Message $logMessage -LogFile $LogFile
            continue
        }
        
        $archiveRoot = Get-ArchiveFolder $account
        if (-not $archiveRoot) {
            $logMessage = "[$($account.Name)] No 'Archive' folder found, skipping."
            Write-Log -Message $logMessage -LogFile $LogFile
            continue
        }

        $year = $Today.ToString("yyyy")
        $month = $Today.ToString("yyyy-MM")

        # Ensure year folder exists or create in live mode
        $yearFolder = $archiveRoot.Folders | Where-Object { $_.Name -eq $year }
        if (-not $yearFolder -and -not $DryRun) {
            $archiveRoot.Folders.Add($year) | Out-Null
            $yearFolder = $archiveRoot.Folders | Where-Object { $_.Name -eq $year }
        }

        # Ensure month folder exists or create in live mode
        $monthFolder = $yearFolder.Folders | Where-Object { $_.Name -eq $month }
        if (-not $monthFolder -and -not $DryRun) {
            $yearFolder.Folders.Add($month) | Out-Null
            $monthFolder = $yearFolder.Folders | Where-Object { $_.Name -eq $month }
        }

        # Safe Inbox retrieval
        $inbox = $null
        try { $inbox = $account.Folders.Item("Inbox") } catch {}
        if (-not $inbox) {
            $logMessage = "[$($account.Name)] No Inbox folder, skipping message scan."
            Write-Log -Message $logMessage -LogFile $LogFile
            continue
        }

        # Get static array of MailItems
        $rawItems = @()
        try {
            $rawItems = @($inbox.Items | Where-Object { $_.Class -eq 43 })
        }
        catch {
            $logMessage = "[$($account.Name)] Could not retrieve mail items: $_"
            Write-Log -Message $logMessage -LogFile $LogFile
            continue
        }

        if ($rawItems.Count -eq 0) {
            $logMessage = "[$($account.Name)] No messages found to process."
            Write-Log -Message $logMessage -LogFile $LogFile
            continue
        }

        # Deduplicate by Subject+DateTime composite key, then sort
        $seenKeys = @{}
        $deduped = foreach ($mail in $rawItems) {
            $key = "$($mail.Subject)|$($mail.ReceivedTime.ToString('o'))"
            if (-not $seenKeys.ContainsKey($key)) {
                $seenKeys[$key] = $true
                $mail
            }
        }
        $sortedItems = $deduped | Sort-Object ReceivedTime

        $emailCount = 0
        foreach ($mail in $sortedItems) {
            $emailCount++
            
            # Limit to 100 emails per mailbox during dry-run for faster testing
            if ($DryRun -and $emailCount -gt 100) {
                $limitMessage = "[$($account.Name)] Reached 100 email limit for testing (dry-run mode)"
                Write-Log -Message $limitMessage -LogFile $LogFile
                Write-Host "  Reached 100 email limit for testing (dry-run mode)" -ForegroundColor Yellow
                break
            }

            # Apply skip rules from config
            $skipMatch = $false
            foreach ($rule in $SkipRules) {
                if ($account.Name -eq $rule.Mailbox) {
                    foreach ($subj in $rule.Subjects) {
                        if ($mail.Subject -match [regex]::Escape($subj)) {
                            $skipMessage = "[$($account.Name)] SKIP: $($mail.ReceivedTime.ToString('yyyy-MM-dd')) : $($mail.Subject)"
                            Write-Log -Message $skipMessage -LogFile $LogFile
                            $skipMatch = $true
                            break
                        }
                    }
                }
                if ($skipMatch) { break }
            }
            if ($skipMatch) { continue }

            if ($mail.ReceivedTime -lt $CutOff) {
                $logEntry = "[$($account.Name)] $($mail.ReceivedTime.ToString('yyyy-MM-dd')) : $($mail.Subject)"
                if ($DryRun) {
                    $dryRunMessage = "DRY-RUN: Would move -> $logEntry"
                    Write-Log -Message $dryRunMessage -LogFile $LogFile
                }
                else {
                    $mail.Move($monthFolder) | Out-Null
                    $movedMessage = "MOVED: $logEntry"
                    Write-Log -Message $movedMessage -LogFile $LogFile
                }
            }
        }

    }
    catch {
        $errorMessage = "[$($account.Name)] Error: $_"
        Write-Log -Message $errorMessage -LogFile $LogFile
    }
}

$completionMessage = "=== Completed at $(Get-Date) ==="
Write-Log -Message $completionMessage -LogFile $LogFile

# === Post-Dry-Run User Interaction ===
if ($DryRun) {
    Write-Host ""
    Write-Host "=== Dry-Run Completed Successfully! ===" -ForegroundColor Green
    Write-Host ""
    Write-Host "The dry-run has finished processing your emails." -ForegroundColor White
    Write-Host "Log file created: $LogFile" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Yellow
    Write-Host "1. Check the log file to review what emails would be archived" -ForegroundColor White
    Write-Host "2. Verify the archive folder structure is correct" -ForegroundColor White
    Write-Host "3. Make any adjustments to config.json if needed" -ForegroundColor White
    Write-Host ""
    
    # Ask user if they want to switch to live mode
    Write-Host "Would you like to switch to live mode and run the actual archiving now?" -ForegroundColor Cyan
    Write-Host "This will move the emails that were identified in the dry-run." -ForegroundColor White
    Write-Host ""
    Write-Host "Options:" -ForegroundColor Yellow
    Write-Host "1. Yes - Switch to live mode and archive emails now" -ForegroundColor White
    Write-Host "2. No - Exit and run manually later" -ForegroundColor White
    Write-Host ""
    
    do {
        $liveModeChoice = Read-Host "Enter choice (1-2)"
        if ($liveModeChoice -match '^[1-2]$') {
            break
        }
        Write-Host "Please enter 1 or 2." -ForegroundColor Red
    } while ($true)
    
    if ($liveModeChoice -eq '1') {
        Write-Host ""
        Write-Host "=== Switching to Live Mode ===" -ForegroundColor Cyan
        Write-Host "Updating config.json to set DryRun = false..." -ForegroundColor White
        
        try {
            # Update config to live mode
            $config.DryRun = $false
            $config | ConvertTo-Json -Depth 3 | Out-File $configPath -Encoding UTF8
            Write-Host "✅ Configuration updated to live mode" -ForegroundColor Green
            
            Write-Host ""
            Write-Host "⚠️  WARNING: This will now move emails to the archive folders!" -ForegroundColor Red
            Write-Host "Make sure you've reviewed the dry-run results and are ready to proceed." -ForegroundColor Yellow
            Write-Host ""
            
            $confirmLive = Read-Host "Type 'YES' to confirm you want to proceed with live archiving"
            if ($confirmLive -eq 'YES') {
                Write-Host ""
                Write-Host "=== Starting Live Archive Process ===" -ForegroundColor Green
                Write-Host "Processing emails in live mode..." -ForegroundColor White
                Write-Host ""
                
                # Update variables for live mode
                $DryRun = $false
                
                # Create new log file for live run
                $liveLogFile = Join-Path $LogPath ("ArchiveLog_LIVE_" + $Today.ToString("yyyy-MM-dd_HH-mm-ss") + ".txt")
                "=== Outlook Auto-Archive LIVE RUN ===" | Out-File -FilePath $liveLogFile -Encoding UTF8
                "Retention: $RetentionDays days" | Out-File -FilePath $liveLogFile -Append -Encoding UTF8
                "Cutoff: $CutOff" | Out-File -FilePath $liveLogFile -Append -Encoding UTF8
                "Started at: $(Get-Date)" | Out-File -FilePath $liveLogFile -Append -Encoding UTF8
                
                Write-Host "Live mode log file: $liveLogFile" -ForegroundColor Cyan
                Write-Host ""
                
                # Re-run the archive process in live mode
                $liveEmailsProcessed = 0
                $liveEmailsMoved = 0
                
                foreach ($account in $namespace.Folders) {
                    try {
                        Write-Host "Processing account: $($account.Name)" -ForegroundColor Cyan
                        
                        # Skip non-email account types
                        $skipAccountTypes = @("Internet Calendars", "SharePoint Lists", "Public Folders", "Calendar", "Contacts", "Tasks", "Notes")
                        if ($skipAccountTypes -contains $account.Name) {
                            $logMessage = "[$($account.Name)] Skipping non-email account type."
                            Write-Log -Message $logMessage -LogFile $liveLogFile
                            continue
                        }
                        
                        $archiveRoot = Get-ArchiveFolder $account
                        if (-not $archiveRoot) {
                            $logMessage = "[$($account.Name)] No 'Archive' folder found, skipping."
                            Write-Log -Message $logMessage -LogFile $liveLogFile
                            continue
                        }

                        $year = $Today.ToString("yyyy")
                        $month = $Today.ToString("yyyy-MM")

                        # Create year and month folders for live mode
                        $yearFolder = $archiveRoot.Folders | Where-Object { $_.Name -eq $year }
                        if (-not $yearFolder) {
                            $archiveRoot.Folders.Add($year) | Out-Null
                            $yearFolder = $archiveRoot.Folders | Where-Object { $_.Name -eq $year }
                            Write-Host "  Created year folder: $year" -ForegroundColor Green
                        }

                        $monthFolder = $yearFolder.Folders | Where-Object { $_.Name -eq $month }
                        if (-not $monthFolder) {
                            $yearFolder.Folders.Add($month) | Out-Null
                            $monthFolder = $yearFolder.Folders | Where-Object { $_.Name -eq $month }
                            Write-Host "  Created month folder: $month" -ForegroundColor Green
                        }

                        # Safe Inbox retrieval
                        $inbox = $null
                        try { $inbox = $account.Folders.Item("Inbox") } catch {}
                        if (-not $inbox) {
                            $logMessage = "[$($account.Name)] No Inbox folder, skipping message scan."
                            Write-Log -Message $logMessage -LogFile $liveLogFile
                            continue
                        }

                        # Get static array of MailItems
                        $rawItems = @()
                        try {
                            $rawItems = @($inbox.Items | Where-Object { $_.Class -eq 43 })
                        }
                        catch {
                            $logMessage = "[$($account.Name)] Could not retrieve mail items: $_"
                            Write-Log -Message $logMessage -LogFile $liveLogFile
                            continue
                        }

                        if ($rawItems.Count -eq 0) {
                            $logMessage = "[$($account.Name)] No messages found to process."
                            Write-Log -Message $logMessage -LogFile $liveLogFile
                            continue
                        }

                        # Deduplicate by Subject+DateTime composite key, then sort
                        $seenKeys = @{}
                        $deduped = foreach ($mail in $rawItems) {
                            $key = "$($mail.Subject)|$($mail.ReceivedTime.ToString('o'))"
                            if (-not $seenKeys.ContainsKey($key)) {
                                $seenKeys[$key] = $true
                                $mail
                            }
                        }
                        $sortedItems = $deduped | Sort-Object ReceivedTime

                        foreach ($mail in $sortedItems) {
                            $liveEmailsProcessed++

                            # Apply skip rules from config
                            $skipMatch = $false
                            foreach ($rule in $SkipRules) {
                                if ($account.Name -eq $rule.Mailbox) {
                                    foreach ($subj in $rule.Subjects) {
                                        if ($mail.Subject -match [regex]::Escape($subj)) {
                                            $skipMessage = "[$($account.Name)] SKIP: $($mail.ReceivedTime.ToString('yyyy-MM-dd')) : $($mail.Subject)"
                                            Write-Log -Message $skipMessage -LogFile $liveLogFile
                                            $skipMatch = $true
                                            break
                                        }
                                    }
                                }
                                if ($skipMatch) { break }
                            }
                            if ($skipMatch) { continue }

                            if ($mail.ReceivedTime -lt $CutOff) {
                                try {
                                    $mail.Move($monthFolder) | Out-Null
                                    $movedMessage = "MOVED: [$($account.Name)] $($mail.ReceivedTime.ToString('yyyy-MM-dd')) : $($mail.Subject)"
                                    Write-Log -Message $movedMessage -LogFile $liveLogFile
                                    $liveEmailsMoved++
                                    
                                    # Show progress every 10 emails
                                    if ($liveEmailsMoved % 10 -eq 0) {
                                        Write-Host "  Moved $liveEmailsMoved emails so far..." -ForegroundColor Green
                                    }
                                }
                                catch {
                                    $errorMessage = "ERROR MOVING: [$($account.Name)] $($mail.ReceivedTime.ToString('yyyy-MM-dd')) : $($mail.Subject) - $_"
                                    Write-Log -Message $errorMessage -LogFile $liveLogFile
                                    Write-Host "  ❌ Error moving email: $_" -ForegroundColor Red
                                }
                            }
                        }

                    }
                    catch {
                        $errorMessage = "[$($account.Name)] Error: $_"
                        Write-Log -Message $errorMessage -LogFile $liveLogFile
                    }
                }
                
                $liveCompletionMessage = "=== Live Archive Completed at $(Get-Date) ==="
                Write-Log -Message $liveCompletionMessage -LogFile $liveLogFile
                Write-Log -Message "Total emails processed: $liveEmailsProcessed" -LogFile $liveLogFile
                Write-Log -Message "Total emails moved: $liveEmailsMoved" -LogFile $liveLogFile
                
                Write-Host ""
                Write-Host "=== Live Archive Completed Successfully! ===" -ForegroundColor Green
                Write-Host "Total emails processed: $liveEmailsProcessed" -ForegroundColor White
                Write-Host "Total emails moved: $liveEmailsMoved" -ForegroundColor Green
                Write-Host "Live mode log file: $liveLogFile" -ForegroundColor Cyan
                Write-Host ""
                Write-Host "✅ Your emails have been successfully archived!" -ForegroundColor Green
                Write-Host "You can find them in your Outlook archive folders organized by year/month." -ForegroundColor White
                
            }
            else {
                Write-Host ""
                Write-Host "Live archiving cancelled by user." -ForegroundColor Yellow
                Write-Host "You can run the script again later when you're ready." -ForegroundColor White
            }
            
        }
        catch {
            Write-Host "❌ Error updating configuration: $_" -ForegroundColor Red
            Write-Host "You can manually edit config.json and set 'DryRun': false" -ForegroundColor Yellow
        }
    }
    else {
        Write-Host ""
        Write-Host "Exiting without switching to live mode." -ForegroundColor Yellow
        Write-Host "To run in live mode later:" -ForegroundColor White
        Write-Host "1. Edit config.json and set 'DryRun': false" -ForegroundColor Gray
        Write-Host "2. Run the script again" -ForegroundColor Gray
    }
}
