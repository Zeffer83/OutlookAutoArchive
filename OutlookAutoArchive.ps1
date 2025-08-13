<#
.SYNOPSIS
  Auto-archive Outlook emails with options from config.json
#>

# Version: 1.7.0
# Author: Ryan Zeffiretti
# Description: Auto-archive Outlook emails with options from config.json

Add-Type -AssemblyName Microsoft.Office.Interop.Outlook
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

# === Load config ===
# Handle path for both script and executable
if ($PSScriptRoot) {
    $scriptDir = $PSScriptRoot
} else {
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
            RetentionDays = 14
            DryRun        = $true
            LogPath       = "%USERPROFILE%\Documents\OutlookAutoArchiveLogs"
            GmailLabel    = "OutlookArchive"
            OnFirstRun    = $true
            ArchiveFolders = @{}
            SkipRules     = @(
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
        Add-Type -AssemblyName Microsoft.Office.Interop.Outlook
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        Write-Host "✅ Connected to Outlook" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Failed to connect to Outlook: $_" -ForegroundColor Red
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
                            $newLabel = $account.Folders.Add($gmailLabel)
                            Write-Host "  ✅ Created Gmail label '$gmailLabel'" -ForegroundColor Green
                            $foldersCreated++
                            # Store the Gmail label path in config
                            $config.ArchiveFolders[$account.Name] = "GmailLabel:$gmailLabel"
                        }
                        catch {
                            Write-Host "  ❌ Failed to create Gmail label: $_" -ForegroundColor Red
                            $errors++
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
                $archivePath = ""
                
                # Check root level first
                try {
                    $archiveFolder = $account.Folders.Item("Archive")
                    $archivePath = "Archive"
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
                            $archivePath = "Inbox\Archive"
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
    
    # Initialize ArchiveFolders if it doesn't exist
    if (-not $config.ArchiveFolders) {
        $config.ArchiveFolders = @{}
    }
    
    # Save updated config
    try {
        $config | ConvertTo-Json -Depth 3 | Out-File $configPath -Encoding UTF8
        Write-Host "✅ Configuration saved" -ForegroundColor Green
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
     Write-Host "2. When Outlook starts (recommended)" -ForegroundColor White
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
         
         # Create daily scheduled task
         try {
             $taskName = "Outlook Auto Archive"
             $taskDescription = "Automatically archive old emails from Outlook"
             $scriptPath = if ($PSScriptRoot) { Join-Path $PSScriptRoot "OutlookAutoArchive.exe" } else { Join-Path (Get-Location) "OutlookAutoArchive.exe" }
             
             # Check if executable exists, fall back to PowerShell script
             if (-not (Test-Path $scriptPath)) {
                 $scriptPath = if ($PSScriptRoot) { Join-Path $PSScriptRoot "OutlookAutoArchive.ps1" } else { Join-Path (Get-Location) "OutlookAutoArchive.ps1" }
                 $arguments = "-ExecutionPolicy Bypass -File `"$scriptPath`""
                 $program = "powershell.exe"
             } else {
                 $arguments = ""
                 $program = $scriptPath
             }
             
             # Create the scheduled task
             $createTaskCmd = "schtasks /create /tn `"$taskName`" /tr `"$program`""
             if ($arguments) { $createTaskCmd += " /sc daily /st $scheduledTime /f" } else { $createTaskCmd += " /sc daily /st $scheduledTime /f" }
             
             Write-Host "Creating scheduled task..." -ForegroundColor Yellow
             $result = Invoke-Expression $createTaskCmd
             
             if ($LASTEXITCODE -eq 0) {
                 Write-Host "✅ Daily scheduled task created successfully!" -ForegroundColor Green
                 Write-Host "Task will run daily at $scheduledTime" -ForegroundColor White
             } else {
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
     elseif ($scheduleChoice -eq '2') {
         Write-Host ""
         Write-Host "Setting up Outlook startup task..." -ForegroundColor Cyan
         
         try {
             # Check if Setup_OutlookStartup_Task.ps1 exists
             $setupScriptPath = if ($PSScriptRoot) { Join-Path $PSScriptRoot "Setup_OutlookStartup_Task.ps1" } else { Join-Path (Get-Location) "Setup_OutlookStartup_Task.ps1" }
             
             if (Test-Path $setupScriptPath) {
                 Write-Host "Running Outlook startup task setup..." -ForegroundColor Yellow
                 & powershell.exe -ExecutionPolicy Bypass -File $setupScriptPath
                 Write-Host "✅ Outlook startup task setup completed!" -ForegroundColor Green
             } else {
                 Write-Host "⚠️  Setup script not found. Creating basic startup task..." -ForegroundColor Yellow
                 
                 # Create a basic startup task
                 $taskName = "Outlook Auto Archive - Startup"
                 $scriptPath = if ($PSScriptRoot) { Join-Path $PSScriptRoot "OutlookAutoArchive.exe" } else { Join-Path (Get-Location) "OutlookAutoArchive.exe" }
                 
                 if (-not (Test-Path $scriptPath)) {
                     $scriptPath = if ($PSScriptRoot) { Join-Path $PSScriptRoot "OutlookAutoArchive.ps1" } else { Join-Path (Get-Location) "OutlookAutoArchive.ps1" }
                     $arguments = "-ExecutionPolicy Bypass -File `"$scriptPath`""
                     $program = "powershell.exe"
                 } else {
                     $arguments = ""
                     $program = $scriptPath
                 }
                 
                 $createTaskCmd = "schtasks /create /tn `"$taskName`" /tr `"$program`" /sc onstart /delay 0000:30 /f"
                 if ($arguments) { $createTaskCmd += " $arguments" }
                 
                 $result = Invoke-Expression $createTaskCmd
                 
                 if ($LASTEXITCODE -eq 0) {
                     Write-Host "✅ Startup task created successfully!" -ForegroundColor Green
                     Write-Host "Task will run 30 seconds after system startup" -ForegroundColor White
                 } else {
                     Write-Host "⚠️  Could not create startup task automatically." -ForegroundColor Yellow
                 }
             }
         }
         catch {
             Write-Host "❌ Error creating startup task: $_" -ForegroundColor Red
             Write-Host "You can set up scheduling manually later." -ForegroundColor Yellow
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
}

# === Apply config settings ===
$RetentionDays = [int]$config.RetentionDays
$DryRun = [bool]$config.DryRun

# Process log path with proper error handling
$rawLogPath = $config.LogPath
if ([string]::IsNullOrEmpty($rawLogPath)) {
    $rawLogPath = "%USERPROFILE%\Documents\OutlookAutoArchiveLogs"
    Write-Host "LogPath was empty, using default: $rawLogPath"
}

# Handle both escaped and unescaped backslashes
$LogPath = $rawLogPath -replace '%USERPROFILE%', $env:USERPROFILE
$LogPath = $LogPath -replace '\\\\', '\'  # Fix double backslashes

if ([string]::IsNullOrEmpty($LogPath)) {
    $LogPath = "$env:USERPROFILE\Documents\OutlookAutoArchiveLogs"
    Write-Host "LogPath processing failed, using fallback: $LogPath"
}

Write-Host "Using log path: $LogPath"

$Today = Get-Date
$CutOff = $Today.AddDays(-$RetentionDays)
$GmailLabel = $config.GmailLabel
$SkipRules = $config.SkipRules

# === Check if Outlook is running ===
try {
    $outlookProcesses = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
    if (-not $outlookProcesses) {
        Write-Host "Outlook is not running. Please start Outlook and try again."
        Write-Host "The script requires Outlook to be running to access email data."
        exit 1
    }
    Write-Host "Outlook is running. Proceeding with archive process..."
}
catch {
    Write-Host "Could not check Outlook status. Proceeding anyway..."
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
        $result = New-Item -Path $LogPath -ItemType Directory -Force -ErrorAction Stop
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
    if ($config.ArchiveFolders -and $config.ArchiveFolders.ContainsKey($account.Name)) {
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
                } catch {}
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

        foreach ($mail in $sortedItems) {

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
