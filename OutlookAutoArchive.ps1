<#
.SYNOPSIS
    Auto-archive Outlook emails with options from config.json
    
.DESCRIPTION
    This PowerShell script automatically archives old emails from Outlook accounts based on configurable retention periods.
    It supports both regular email accounts (using folders) and Gmail accounts (using labels).
    
    Key Features:
    - Configurable retention periods (default: 14 days)
    - Dry-run mode for testing before actual archiving
    - Support for Gmail labels and regular email folders
    - Automatic folder creation (year/month structure)
    - Skip rules for specific subjects or mailboxes
    - Comprehensive logging
    - Windows Task Scheduler integration
    - First-run setup wizard
    
.PARAMETER None
    This script uses configuration from config.json file
    
.EXAMPLE
    .\OutlookAutoArchive.ps1
    Runs the script with settings from config.json
    
.NOTES
                   Version: 2.9.5
    Author: Ryan Zeffiretti
    License: MIT
    Requires: Microsoft Outlook to be running
    Requires: PowerShell 5.1 or later
    
    Installation:
    - First run installs to C:\Users\$env:USERNAME\OutlookAutoArchive
    - Creates config.json with default settings
    - Sets up archive folders/labels for each email account
    
    Configuration:
    - Edit config.json to change retention days, dry-run mode, etc.
    - Set 'DryRun': false when ready for live archiving
    - Logs are stored in the Logs folder within installation directory
#>

# Version: 2.9.5
# Author: Ryan Zeffiretti
# Description: Auto-archive Outlook emails with options from config.json
# License: MIT
# Last Updated: 2025-08-14

# =============================================================================
# OUTLOOK INTEROP ASSEMBLY LOADING
# =============================================================================
# Try to load Outlook Interop assembly, but don't fail if it's not available
# This provides better type safety and IntelliSense support when available
# If not available, we fall back to COM objects which work in all environments
try {
    Add-Type -AssemblyName Microsoft.Office.Interop.Outlook -ErrorAction SilentlyContinue
    Write-Host "[OK] Outlook Interop assembly loaded successfully" -ForegroundColor Green
}
catch {
    Write-Host "Note: Microsoft.Office.Interop.Outlook assembly not found, will use COM objects directly" -ForegroundColor Yellow
}

# Initialize Outlook COM objects (will be set up later when needed)
# These are global variables that will be populated when we connect to Outlook
$outlook = $null      # Main Outlook application object
$namespace = $null    # MAPI namespace for accessing folders and accounts

# =============================================================================
# WINDOWS SECURITY - EXECUTABLE UNBLOCKING
# =============================================================================
# Windows automatically blocks executables downloaded from the internet for security
# This section detects if the executable is blocked and attempts to unblock it
# This is essential for user experience when downloading from GitHub or other sources
try {
    # Determine the path to the executable (works for both .ps1 and compiled .exe)
    $currentExePath = if ($PSScriptRoot) { Join-Path $PSScriptRoot "OutlookAutoArchive.exe" } else { Join-Path (Get-Location) "OutlookAutoArchive.exe" }
    
    if (Test-Path $currentExePath) {
        # Check for Zone.Identifier alternate data stream (indicates internet download)
        $zoneInfo = Get-ItemProperty -Path $currentExePath -Name Zone.Identifier -ErrorAction SilentlyContinue
        if ($zoneInfo -and $zoneInfo.'Zone.Identifier') {
            Write-Host ""
            Write-Host "[!] Windows has blocked this executable because it was downloaded from the internet." -ForegroundColor Yellow
            Write-Host "Attempting to unblock the file automatically..." -ForegroundColor Cyan
            
            try {
                # Use PowerShell's Unblock-File cmdlet to remove the block
                Unblock-File -Path $currentExePath -ErrorAction Stop
                Write-Host "[OK] Successfully unblocked the executable!" -ForegroundColor Green
                Write-Host "You can now run the application normally." -ForegroundColor White
            }
            catch {
                # If automatic unblocking fails, provide manual instructions
                Write-Host "[ERROR] Could not automatically unblock the file." -ForegroundColor Red
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

# =============================================================================
# CONFIGURATION LOADING AND INITIALIZATION
# =============================================================================
# This section handles loading the configuration file and setting up default values
# The config.json file stores user preferences and discovered archive folder paths

# Determine the installation directory (hardcoded for consistency)
$installPath = "C:\Users\$env:USERNAME\OutlookAutoArchive"
$configPath = Join-Path $installPath 'config.json'
$scriptDir = $installPath  # Set script directory for normal operation

Write-Host "Installation directory: $installPath"
Write-Host "Config path: $configPath"

# Initialize config object (will be populated from file or defaults)
$config = $null

# Try to load existing configuration file
if (Test-Path $configPath) {
    try {
        # Read and parse the JSON configuration file
        $config = Get-Content $configPath -Raw | ConvertFrom-Json
        Write-Host "[OK] Loaded existing configuration" -ForegroundColor Green
    }
    catch {
        # Handle corrupted or invalid JSON files
        Write-Host "[!] Invalid JSON in config.json, will create new configuration" -ForegroundColor Yellow
        $config = $null
    }
}

# If no config exists or is invalid, initialize with safe default values
if (-not $config) {
    Write-Host "No valid configuration found. Initializing with default settings." -ForegroundColor Cyan
    $config = @{
        RetentionDays      = 14                    # Days to keep emails in Inbox before archiving
        DryRun             = $true                 # Safety mode - don't actually move emails
        LogPath            = "./Logs"              # Relative path for log files
        GmailLabel         = "OutlookArchive"      # Label name for Gmail accounts
        OnFirstRun         = $true                 # Flag to trigger first-run setup wizard
        ArchiveFolders     = @{}                   # Hash table of discovered archive folder paths
        MonitoringInterval = 4                     # Hours between continuous monitoring runs
        SkipRules          = @(                    # Rules to skip archiving specific emails
            @{
                Mailbox  = "Your Mailbox Name"     # Example skip rule
                Subjects = @("Subject Pattern 1", "Subject Pattern 2")
            }
        )
    }
}

# =============================================================================
# FIRST RUN SETUP WIZARD
# =============================================================================
# This section handles the initial setup when the script is run for the first time
# It guides users through configuration, creates archive folders, and sets up scheduling
if ($config.OnFirstRun -eq $true) {
    # ASCII Art Banner
    Write-Host ""
    Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║                                                              ║" -ForegroundColor Cyan
    Write-Host "║   OUTLOOK AUTO ARCHIVE - FIRST RUN                           ║" -ForegroundColor Cyan
    Write-Host "║   Welcome to Your Email Archiving Setup                      ║" -ForegroundColor Cyan
    Write-Host "║                                                              ║" -ForegroundColor Cyan
    Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "[TARGET] This appears to be your first time running the script." -ForegroundColor White
    Write-Host "[STEPS] Let's set up your archive folders and configuration." -ForegroundColor White
    Write-Host ""
    
    # =================================================================
    # ADMIN RIGHTS CHECK FOR SCHEDULED TASK CREATION
    # =================================================================
    # Check admin rights early for scheduling setup
    # Windows Task Scheduler requires admin privileges to create tasks
    Write-Host "[SEARCH] Checking system requirements..." -ForegroundColor Cyan
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    
    if (-not $isAdmin) {
        Write-Host "[!]  Note: You're not running as Administrator" -ForegroundColor Yellow
        Write-Host "This is fine for normal usage, but you'll need admin rights for scheduled task creation." -ForegroundColor White
        Write-Host "You can:" -ForegroundColor Cyan
        Write-Host "1. Continue with setup (you can set up scheduling later with admin rights)" -ForegroundColor White
        Write-Host "2. Restart as Administrator now" -ForegroundColor White
        Write-Host ""
        
        # Get user choice for admin rights handling
        do {
            $adminChoice = Read-Host "Continue with setup or restart as Administrator? (1/2)"
            if ($adminChoice -match '^[1-2]$') {
                break
            }
            Write-Host "Please enter 1 or 2." -ForegroundColor Red
        } while ($true)
        
        if ($adminChoice -eq '2') {
            # Restart the script with elevated privileges
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
        Write-Host "[OK] Running with Administrator privileges - all features available" -ForegroundColor Green
        Write-Host ""
    }
    
    # =================================================================
    # INSTALLATION DIRECTORY SETUP
    # =================================================================
    # Set installation location to user's home directory for consistency
    # This ensures the app always knows where it's installed
    $currentLocation = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }
    $installPath = "C:\Users\$env:USERNAME\OutlookAutoArchive"
    
    Write-Host "Installation location: $installPath" -ForegroundColor Green
    
    # Check if we need to move files from current location to installation directory
    if ($installPath -ne $currentLocation) {
        Write-Host ""
        Write-Host "Setting up installation at: $installPath" -ForegroundColor Cyan
        
        try {
            # Create the installation directory if it doesn't exist
            if (-not (Test-Path $installPath)) {
                New-Item -Path $installPath -ItemType Directory -Force | Out-Null
                Write-Host "[OK] Created installation directory" -ForegroundColor Green
            }
            
            # Copy only essential files to the installation location
            # We only copy the executable - config.json will be created during setup
            $filesToCopy = @(
                "OutlookAutoArchive.exe"    # Main application executable
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
            
            Write-Host "[OK] Copied $filesCopied files to installation directory" -ForegroundColor Green
            
            # Create a simple README.txt for users with essential information
            # This provides users with quick start instructions and troubleshooting help
            $readmeContent = @"
 Outlook Auto Archive - Version 2.9.0
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

 Version 2.9.0 - Fixed LogPath handling and improved path processing
"@
            
            $readmePath = Join-Path $installPath "README.txt"
            $readmeContent | Out-File -FilePath $readmePath -Encoding UTF8
            Write-Host "[OK] Created user-friendly README.txt" -ForegroundColor Green
            
            # Update the script directory for the rest of the setup
            $scriptDir = $installPath
            $configPath = Join-Path $installPath 'config.json'
            
            Write-Host ""
            Write-Host "Installation completed successfully!" -ForegroundColor Green
            Write-Host "The application is now installed at: $installPath" -ForegroundColor White
            Write-Host ""
            Write-Host "Note: You can now delete the original files from: $currentLocation" -ForegroundColor Yellow
            Write-Host "The application will run from the new location." -ForegroundColor White
            Write-Host ""
            
        }
        catch {
            Write-Host "[ERROR] Error during installation: $_" -ForegroundColor Red
            Write-Host "Continuing with installation directory..." -ForegroundColor Yellow
            $scriptDir = $installPath
            $configPath = Join-Path $installPath 'config.json'
        }
    }
    else {
        # Use installation directory
        $scriptDir = $installPath
        $configPath = Join-Path $installPath 'config.json'
    }
    
    # =================================================================
    # OUTLOOK CONNECTION AND VALIDATION
    # =================================================================
    # Check if Outlook is running before attempting to connect
    # This prevents errors and provides clear user feedback
    try {
        $outlookProcesses = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
        if (-not $outlookProcesses) {
            Write-Host "[ERROR] Outlook is not running. Please start Outlook and run the script again." -ForegroundColor Red
            Write-Host "The setup requires Outlook to be running to access your email accounts." -ForegroundColor Yellow
            exit 1
        }
        Write-Host "[OK] Outlook is running" -ForegroundColor Green
    }
    catch {
        Write-Host "[!]  Could not check Outlook status. Proceeding anyway..." -ForegroundColor Yellow
    }
    
    # Connect to Outlook using COM objects
    # This establishes the connection needed to access email accounts and folders
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        Write-Host "[OK] Connected to Outlook" -ForegroundColor Green
    }
    catch {
        Write-Host "[ERROR] Failed to connect to Outlook: $_" -ForegroundColor Red
        Write-Host "Make sure Outlook is running and you have the necessary permissions." -ForegroundColor Yellow
        exit 1
    }
    
    # =================================================================
    # EMAIL ACCOUNT DISCOVERY AND CLASSIFICATION
    # =================================================================
    # Scan all Outlook accounts and classify them as Gmail or regular accounts
    # This helps determine the appropriate archiving method (labels vs folders)
    Write-Host ""
    Write-Host "[EMAIL] Scanning your email accounts..." -ForegroundColor Cyan
    
    $accounts = @()           # All discovered accounts
    $gmailAccounts = @()      # Gmail accounts (use labels)
    $regularAccounts = @()    # Regular email accounts (use folders)
    
    foreach ($account in $namespace.Folders) {
        $accounts += $account.Name
        
        # Check if this looks like a Gmail account based on email domain
        $isGmail = $account.Name -like "*@gmail.com" -or $account.Name -like "*@googlemail.com" -or $account.Name -like "*@gmail.co.uk"
        
        if ($isGmail) {
            $gmailAccounts += $account.Name
        }
        else {
            $regularAccounts += $account.Name
        }
    }
    
    Write-Host "[OK] Found $($accounts.Count) email account(s):" -ForegroundColor Green
    foreach ($account in $accounts) {
        Write-Host "  [EMAIL] $account" -ForegroundColor White
    }
    
    if ($gmailAccounts.Count -gt 0) {
        Write-Host ""
        Write-Host "[SEARCH] Gmail accounts detected: $($gmailAccounts -join ', ')" -ForegroundColor Yellow
        Write-Host "[TIP] Note: Gmail accounts will use labels instead of folders for archiving." -ForegroundColor Gray
    }
    
    Write-Host ""
    
    # =================================================================
    # RETENTION PERIOD CONFIGURATION
    # =================================================================
    # Get user preference for how long emails should stay in Inbox before archiving
    # This is a critical setting that affects which emails get moved
    Write-Host ""
    Write-Host "[TIME] RETENTION PERIOD CONFIGURATION:" -ForegroundColor Yellow
    Write-Host "How many days should emails stay in your Inbox before being archived?" -ForegroundColor Cyan
    Write-Host "[TIP] Recommended: 14-30 days" -ForegroundColor Gray
    Write-Host ""
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
        Write-Host "[ERROR] Please enter a valid positive number." -ForegroundColor Red
    } while ($true)
    
    Write-Host "[OK] Retention period set to $retentionDays days" -ForegroundColor Green
    
    # =================================================================
    # GMAIL LABEL CONFIGURATION
    # =================================================================
    # Configure custom label name for Gmail accounts
    # Gmail doesn't allow "Archive" as a label name, so we use a custom name
    $gmailLabel = "OutlookArchive"
    if ($gmailAccounts.Count -gt 0) {
        Write-Host ""
        Write-Host "[LABEL]  GMAIL LABEL CONFIGURATION:" -ForegroundColor Yellow
        Write-Host "For Gmail accounts, what would you like to call your archive label?" -ForegroundColor Cyan
        Write-Host "[!]  Note: 'Archive' is not allowed in Gmail, so we use a custom label name." -ForegroundColor Gray
        Write-Host "[TIP] Recommended: OutlookArchive, MyArchive, or EmailArchive" -ForegroundColor Gray
        Write-Host ""
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
            Write-Host "[ERROR] Please enter a valid label name (letters, numbers, hyphens, underscores only)." -ForegroundColor Red
        } while ($true)
        
        Write-Host "[OK] Gmail archive label set to '$gmailLabel'" -ForegroundColor Green
    }
    
    # =================================================================
    # ARCHIVE FOLDER AND LABEL SETUP
    # =================================================================
    # Check for existing archive folders/labels and create any missing ones
    # This ensures the archiving process has a destination for each account
    Write-Host ""
    Write-Host "[FOLDER] ARCHIVE FOLDER SETUP:" -ForegroundColor Yellow
    Write-Host "Now let's check for existing archive folders and create any missing ones..." -ForegroundColor Cyan
    
    $foldersCreated = 0    # Counter for successfully created folders/labels
    $errors = 0            # Counter for errors encountered during setup
    
    foreach ($account in $namespace.Folders) {
        try {
            Write-Host ""
            Write-Host "Processing account: $($account.Name)" -ForegroundColor Cyan
            
            # Skip non-email account types that can't be archived
            # These account types don't contain emails and would cause errors
            $skipAccountTypes = @("Internet Calendars", "SharePoint Lists", "Public Folders", "Calendar", "Contacts", "Tasks", "Notes")
            if ($skipAccountTypes -contains $account.Name) {
                Write-Host "  [!]  Skipping non-email account type: $($account.Name)" -ForegroundColor Yellow
                continue
            }
            
            # Check if this looks like a Gmail account based on email domain
            $isGmail = $account.Name -like "*@gmail.com" -or $account.Name -like "*@googlemail.com" -or $account.Name -like "*@gmail.co.uk"
            
            if ($isGmail) {
                Write-Host "  Detected Gmail account" -ForegroundColor Gray
                
                # Gmail accounts use labels instead of folders for organization
                # Check if the Gmail label already exists
                $existingLabel = $null
                try {
                    $existingLabel = $account.Folders.Item($gmailLabel)
                }
                catch {}
                
                if ($existingLabel) {
                    Write-Host "  [OK] Gmail label '$gmailLabel' already exists" -ForegroundColor Green
                    # Store the Gmail label path in config for future use
                    $config.ArchiveFolders[$account.Name] = "GmailLabel:$gmailLabel"
                }
                else {
                    Write-Host "  Gmail label '$gmailLabel' not found" -ForegroundColor Yellow
                    $createLabel = Read-Host "  Would you like to create it? (Y/N)"
                    if ($createLabel -eq 'Y' -or $createLabel -eq 'y') {
                        try {
                            # Create the Gmail label
                            $account.Folders.Add($gmailLabel)
                            Write-Host "  [OK] Created Gmail label '$gmailLabel'" -ForegroundColor Green
                            $foldersCreated++
                            # Store the Gmail label path in config for future use
                            $config.ArchiveFolders[$account.Name] = "GmailLabel:$gmailLabel"
                        }
                        catch {
                            # Gmail label creation can sometimes throw errors but still succeed
                            # Check if the label was actually created despite the error
                            try {
                                $testLabel = $account.Folders.Item($gmailLabel)
                                if ($testLabel) {
                                    Write-Host "  [OK] Gmail label '$gmailLabel' was created successfully" -ForegroundColor Green
                                    $foldersCreated++
                                    # Store the Gmail label path in config for future use
                                    $config.ArchiveFolders[$account.Name] = "GmailLabel:$gmailLabel"
                                }
                            }
                            catch {
                                Write-Host "  [!]  Gmail label creation encountered an issue, but this is often normal for Gmail accounts" -ForegroundColor Yellow
                                Write-Host "  The label may still be available in Outlook. You can check manually or try again later." -ForegroundColor Gray
                                $errors++
                            }
                        }
                    }
                    else {
                        Write-Host "  [!]  Skipped creating Gmail label" -ForegroundColor Yellow
                    }
                }
            }
            else {
                Write-Host "  Detected regular email account" -ForegroundColor Gray
                
                # Regular email accounts use folders for organization
                # Check for existing archive folder in common locations
                $archiveFolder = $null
                
                # Check root level first (most common location)
                try {
                    $archiveFolder = $account.Folders.Item("Archive")
                    Write-Host "  [OK] Archive folder already exists at root level" -ForegroundColor Green
                    # Store the archive folder path in config for future use
                    $config.ArchiveFolders[$account.Name] = "Root:Archive"
                }
                catch {
                    # Check Inbox\Archive as alternative location
                    try {
                        $inbox = $account.Folders.Item("Inbox")
                        if ($inbox) {
                            $archiveFolder = $inbox.Folders.Item("Archive")
                            Write-Host "  [OK] Archive folder already exists in Inbox" -ForegroundColor Green
                            # Store the archive folder path in config for future use
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
                        
                        # Get user preference for archive folder location
                        do {
                            $locationChoice = Read-Host "  Enter choice (1-3)"
                            if ($locationChoice -match '^[1-3]$') {
                                break
                            }
                            Write-Host "  Please enter 1, 2, or 3." -ForegroundColor Red
                        } while ($true)
                        
                        if ($locationChoice -eq '1') {
                            # Create archive folder at root level (recommended)
                            try {
                                $archiveFolder = $account.Folders.Add("Archive")
                                Write-Host "  [OK] Created Archive folder at root level" -ForegroundColor Green
                                $foldersCreated++
                                # Store the archive folder path in config for future use
                                $config.ArchiveFolders[$account.Name] = "Root:Archive"
                            }
                            catch {
                                Write-Host "  [ERROR] Failed to create Archive folder: $_" -ForegroundColor Red
                                $errors++
                            }
                        }
                        elseif ($locationChoice -eq '2') {
                            # Create archive folder inside Inbox folder
                            try {
                                $inbox = $account.Folders.Item("Inbox")
                                if ($inbox) {
                                    $archiveFolder = $inbox.Folders.Add("Archive")
                                    Write-Host "  [OK] Created Archive folder in Inbox" -ForegroundColor Green
                                    $foldersCreated++
                                    # Store the archive folder path in config for future use
                                    $config.ArchiveFolders[$account.Name] = "Inbox:Archive"
                                }
                                else {
                                    Write-Host "  [ERROR] Could not access Inbox folder" -ForegroundColor Red
                                    $errors++
                                }
                            }
                            catch {
                                Write-Host "  [ERROR] Failed to create Archive folder: $_" -ForegroundColor Red
                                $errors++
                            }
                        }
                        else {
                            Write-Host "  [!]  Skipped creating Archive folder" -ForegroundColor Yellow
                        }
                    }
                }
            }
        }
        catch {
            Write-Host "  [ERROR] Error processing account '$($account.Name)': $_" -ForegroundColor Red
            $errors++
        }
    }
    
    # =================================================================
    # SETUP SUMMARY AND CONFIGURATION SAVING
    # =================================================================
    # Display summary of what was accomplished during setup
    Write-Host ""
    Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║   SETUP SUMMARY                                              ║" -ForegroundColor Green
    Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Green
    Write-Host ""
    Write-Host "[EMAIL] Accounts processed: $($accounts.Count)" -ForegroundColor White
    Write-Host "[FOLDER] Folders/labels created: $foldersCreated" -ForegroundColor White
    Write-Host "[ERROR] Errors encountered: $errors" -ForegroundColor $(if ($errors -gt 0) { "Red" } else { "Green" })
    
    # Update configuration object with user preferences from setup
    $config.RetentionDays = $retentionDays
    $config.GmailLabel = $gmailLabel
    $config.OnFirstRun = $false  # Mark first run as complete
     
    # Save monitoring interval if it was set during setup
    if ($monitoringInterval) {
        $config.MonitoringInterval = $monitoringInterval
    }
    
    # Ensure installation directory exists before saving config
    if (-not (Test-Path $installPath)) {
        New-Item -Path $installPath -ItemType Directory -Force | Out-Null
        Write-Host "Created installation directory: $installPath" -ForegroundColor Green
    }
    
    # Save updated configuration with discovered archive folder paths
    # This is crucial for future runs to avoid re-scanning for folders
    try {
        $config | ConvertTo-Json -Depth 3 | Out-File $configPath -Encoding UTF8
        Write-Host "[OK] Configuration saved with archive folder paths" -ForegroundColor Green
        Write-Host "Archive folders discovered and stored for future runs:" -ForegroundColor Cyan
        foreach ($accountName in $config.ArchiveFolders.Keys) {
            Write-Host "  - $accountName`: $($config.ArchiveFolders[$accountName])" -ForegroundColor White
        }
    }
    catch {
        Write-Host "[ERROR] Failed to save configuration: $_" -ForegroundColor Red
    }
    
    Write-Host ""
    Write-Host "[SUCCESS] First run setup completed!" -ForegroundColor Green
    Write-Host ""
     
    # =================================================================
    # SCHEDULED TASK SETUP
    # =================================================================
    # Offer to create Windows Task Scheduler tasks for automatic archiving
    # This allows the script to run automatically without user intervention
    Write-Host "[TIME] SCHEDULED TASK SETUP:" -ForegroundColor Yellow
    Write-Host "Would you like to set up automatic scheduling now?" -ForegroundColor Cyan
    Write-Host "This will create a Windows Task Scheduler task to run the archive script automatically." -ForegroundColor White
    Write-Host ""
    Write-Host "[SCHEDULE] Scheduling options:" -ForegroundColor Yellow
    Write-Host "┌─────────────────────────────────────────────────────────────────┐" -ForegroundColor Gray
    Write-Host "│ 1. DAILY ARCHIVING                                              │" -ForegroundColor White
    Write-Host "│    Runs once per day at a specific time (e.g., 2:00 AM)         │" -ForegroundColor Gray
    Write-Host "│    Best for: Users who want predictable, quiet archiving        │" -ForegroundColor Gray
    Write-Host "└─────────────────────────────────────────────────────────────────┘" -ForegroundColor Gray
    Write-Host ""
    Write-Host "┌─────────────────────────────────────────────────────────────────┐" -ForegroundColor Gray
    Write-Host "│ 2. SKIP SCHEDULING FOR NOW                                      │" -ForegroundColor White
    Write-Host "│    You can set up scheduling later using the setup script       │" -ForegroundColor Gray
    Write-Host "└─────────────────────────────────────────────────────────────────┘" -ForegroundColor Gray
    Write-Host ""
     
    do {
        $scheduleChoice = Read-Host "Enter choice (1-2)"
        if ($scheduleChoice -match '^[1-2]$') {
            break
        }
        Write-Host "Please enter 1 or 2." -ForegroundColor Red
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
            Write-Host "[!]  Admin rights required for scheduled task creation" -ForegroundColor Yellow
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
                    Write-Host "[OK] Daily scheduled task created successfully!" -ForegroundColor Green
                    Write-Host "Task will run daily at $scheduledTime" -ForegroundColor White
                }
                else {
                    Write-Host "[!]  Could not create scheduled task automatically." -ForegroundColor Yellow
                    Write-Host "You can create it manually using Task Scheduler:" -ForegroundColor White
                    Write-Host "1. Open Task Scheduler" -ForegroundColor Gray
                    Write-Host "2. Create Basic Task" -ForegroundColor Gray
                    Write-Host "3. Name: Outlook Auto Archive" -ForegroundColor Gray
                    Write-Host "4. Trigger: Daily at $scheduledTime" -ForegroundColor Gray
                    Write-Host "5. Action: Start program: $program" -ForegroundColor Gray
                }
            }
            catch {
                Write-Host "[ERROR] Error creating scheduled task: $_" -ForegroundColor Red
                Write-Host "You can set up scheduling manually later." -ForegroundColor Yellow
            }
        }
    }
    else {
        Write-Host ""
        Write-Host "Scheduling skipped. You can set it up later using:" -ForegroundColor Yellow
        Write-Host "1. Task Scheduler GUI" -ForegroundColor White
        Write-Host "2. setup_task_scheduler.exe" -ForegroundColor White
        Write-Host "3. Manual schtasks command" -ForegroundColor White
    }
     
    Write-Host ""
    Write-Host "[STEPS] NEXT STEPS:" -ForegroundColor Yellow
    Write-Host "1. The script will now run in dry-run mode to test everything" -ForegroundColor White
    Write-Host "2. Check the log files to verify everything works" -ForegroundColor White
    Write-Host "3. When ready, edit config.json and set 'DryRun': false" -ForegroundColor White
    Write-Host "4. Test your scheduled task if you created one" -ForegroundColor White
     
    Write-Host ""
    Write-Host "[!]  IMPORTANT: The dry-run test may take several minutes depending on how many emails you have." -ForegroundColor Yellow
    Write-Host "This is normal - the script is scanning all your emails to show what would be archived." -ForegroundColor White
    Write-Host "Please be patient and don't close the window while it's running." -ForegroundColor White
    
    if ($gmailAccounts.Count -gt 0) {
        Write-Host ""
        Write-Host "[EMAIL] For Gmail users:" -ForegroundColor Cyan
        Write-Host "   • Make sure IMAP is enabled in Gmail settings" -ForegroundColor White
        Write-Host "   • Check 'Show in IMAP' for your labels in Gmail web interface" -ForegroundColor White
        Write-Host "   • It may take a few minutes for labels to sync to Outlook" -ForegroundColor White
    }
    
    Write-Host ""
    Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║   CONTINUING WITH ARCHIVE PROCESS                            ║" -ForegroundColor Cyan
    Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    
    # Reload config to get the updated values from the first-run setup
    try {
        $config = Get-Content $configPath -Raw | ConvertFrom-Json
        Write-Host "[OK] Configuration reloaded with updated settings" -ForegroundColor Green
    }
    catch {
        Write-Host "[!]  Could not reload configuration, continuing with current settings" -ForegroundColor Yellow
    }
}

# =============================================================================
# MAIN PROCESSING - CONFIGURATION APPLICATION
# =============================================================================
# Apply configuration settings to variables used throughout the script
# This section runs after first-run setup (if applicable) or loads from existing config
$RetentionDays = [int]$config.RetentionDays    # Convert to integer for date calculations
$DryRun = [bool]$config.DryRun                 # Convert to boolean for conditional logic

# Process log path with proper error handling
$rawLogPath = $config.LogPath
if ([string]::IsNullOrEmpty($rawLogPath)) {
    $rawLogPath = ".\Logs"
    Write-Host "LogPath was empty, using default: $rawLogPath"
}

# Handle relative paths and environment variables
# First, normalize the path by removing escaped backslashes
$normalizedPath = $rawLogPath -replace '\\\\', '\'  # Fix double backslashes
$normalizedPath = $normalizedPath -replace '//', '/'    # Fix double forward slashes

if ($normalizedPath -like ".\*" -or $normalizedPath -like "./*") {
    # Relative path - make it absolute based on script location
    $LogPath = Join-Path $scriptDir $normalizedPath.Substring(2)
}
else {
    # Handle environment variables for absolute paths
    $LogPath = $normalizedPath -replace '%USERPROFILE%', $env:USERPROFILE
}

# Ensure LogPath is never null or empty
if ([string]::IsNullOrEmpty($LogPath)) {
    $LogPath = Join-Path $scriptDir "Logs"
    Write-Host "LogPath processing failed, using fallback: $LogPath"
}

# Ensure the LogPath is absolute
if (-not [System.IO.Path]::IsPathRooted($LogPath)) {
    $LogPath = Join-Path $scriptDir $LogPath
}

Write-Host "Using log path: $LogPath"

# Calculate date-based variables for archiving logic
$Today = Get-Date
$CutOff = $Today.AddDays(-$RetentionDays)      # Date threshold for archiving
Write-Host "Retention period: $RetentionDays days" -ForegroundColor Cyan
Write-Host "Cutoff date: $CutOff (emails older than this will be archived)" -ForegroundColor Cyan
$GmailLabel = $config.GmailLabel               # Label name for Gmail accounts
$SkipRules = $config.SkipRules                 # Rules for skipping specific emails

# =============================================================================
# OUTLOOK AVAILABILITY CHECK (INTERACTIVE RUNS ONLY)
# =============================================================================
# Check if Outlook is running before attempting to connect
# Note: Scheduled runs handle Outlook availability gracefully in the connection section below
if ([Environment]::UserInteractive) {
    try {
        $outlookProcesses = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
        if (-not $outlookProcesses) {
            Write-Host "[ERROR] Outlook is not running!" -ForegroundColor Red
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
        Write-Host "[OK] Outlook is running. Proceeding with archive process..." -ForegroundColor Green
    }
    catch {
        Write-Host "[!]  Could not check Outlook status. Proceeding anyway..." -ForegroundColor Yellow
    }
}

# =============================================================================
# LOGGING SYSTEM SETUP
# =============================================================================
# Initialize logging system for tracking archive operations
# Logs are essential for troubleshooting and audit trails
$LogFile = $null
try {
    # Ensure LogPath is valid before attempting to create log files
    if ([string]::IsNullOrEmpty($LogPath)) {
        throw "LogPath is null or empty"
    }
    
    # Create log directory if it doesn't exist
    if (-not (Test-Path $LogPath)) { 
        New-Item -Path $LogPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
        Write-Host "Created log directory: $LogPath"
    }
    
    # Create log file path with timestamp for uniqueness
    $LogFile = Join-Path $LogPath ("ArchiveLog_" + $Today.ToString("yyyy-MM-dd_HH-mm-ss") + ".txt")
    
    # Initialize log file with header information
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

# =============================================================================
# OUTLOOK CONNECTION FOR MAIN PROCESSING
# =============================================================================
# Establish connection to Outlook for the main archiving process
# This handles both interactive and scheduled runs with appropriate error handling
if (-not $outlook -or -not $namespace) {
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        Write-Host "[OK] Connected to Outlook for processing" -ForegroundColor Green
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
            Write-Host "[ERROR] Failed to connect to Outlook: $_" -ForegroundColor Red
            Write-Host "Make sure Outlook is running and you have the necessary permissions." -ForegroundColor Yellow
            exit 1
        }
    }
}

# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

# Helper function for safe logging to both console and file
# This ensures consistent logging behavior throughout the script
function Write-Log {
    param(
        [string]$Message,    # Message to log
        [string]$LogFile     # Path to log file (optional)
    )
    
    # Always write to console for immediate feedback
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

# =============================================================================
# ENHANCED LOGGING AND ITEM ANALYSIS FUNCTIONS
# =============================================================================
# These functions provide detailed analysis of items to help identify leftover items

# Global counters for detailed reporting
$global:TotalItemsProcessed = 0
$global:TotalItemsMoved = 0
$global:TotalItemsSkipped = 0
$global:TotalItemsTooRecent = 0
$global:TotalItemsError = 0
$global:ItemTypeBreakdown = @{}

# Helper function to get item type description
function Get-ItemTypeDescription {
    param([int]$Class)
    
    switch ($Class) {
        43 { return "MailItem (Email)" }
        26 { return "MeetingItem (Meeting)" }
        34 { return "PostItem (Post)" }
        35 { return "ContactItem (Contact)" }
        1 { return "AppointmentItem (Calendar)" }
        5 { return "TaskItem (Task)" }
        53 { return "CalendarItem (Calendar Meeting)" }
        default { return "Unknown Class $Class" }
    }
}

# Helper function to log detailed item analysis
function Write-ItemAnalysis {
    param(
        [string]$AccountName,
        [array]$Items,
        [string]$LogFile
    )
    
    if (-not $Items -or $Items.Count -eq 0) {
        Write-Log -Message "[$AccountName] No items found for analysis" -LogFile $LogFile
        return
    }
    
    Write-Log -Message "[$AccountName] === DETAILED ITEM ANALYSIS ===" -LogFile $LogFile
    Write-Log -Message "[$AccountName] Total items found: $($Items.Count)" -LogFile $LogFile
    
    # Group items by class
    $classGroups = $Items | Group-Object Class
    foreach ($group in $classGroups) {
        $className = Get-ItemTypeDescription -Class $group.Name
        Write-Log -Message "[$AccountName] Class $($group.Name) ($className): $($group.Count) items" -LogFile $LogFile
        
        # Update global breakdown
        if (-not $global:ItemTypeBreakdown.ContainsKey($className)) {
            $global:ItemTypeBreakdown[$className] = 0
        }
        $global:ItemTypeBreakdown[$className] += $group.Count
    }
    
    # Check for items older than retention period
    $cutoff = (Get-Date).AddDays(-$RetentionDays)
    $oldItems = $Items | Where-Object { $_.ReceivedTime -and $_.ReceivedTime -lt $cutoff }
    Write-Log -Message "[$AccountName] Items older than $RetentionDays days: $($oldItems.Count)" -LogFile $LogFile
    
    # Check for items older than retention period by type
    $oldMailItems = $oldItems | Where-Object { $_.Class -eq 43 }
    $oldMeetingItems = $oldItems | Where-Object { $_.Class -eq 26 }
    $oldCalendarItems = $oldItems | Where-Object { $_.Class -eq 53 }
    Write-Log -Message "[$AccountName] MailItems older than $RetentionDays days: $($oldMailItems.Count)" -LogFile $LogFile
    Write-Log -Message "[$AccountName] MeetingItems older than $RetentionDays days: $($oldMeetingItems.Count)" -LogFile $LogFile
    Write-Log -Message "[$AccountName] CalendarItems older than $RetentionDays days: $($oldCalendarItems.Count)" -LogFile $LogFile
    
    # Show examples of old items that should be archived
    if ($oldItems.Count -gt 0) {
        Write-Log -Message "[$AccountName] Examples of items that should be archived:" -LogFile $LogFile
        $examples = $oldItems | Select-Object -First 5
        foreach ($item in $examples) {
            try {
                $subject = if ($item.Subject) { $item.Subject.Substring(0, [Math]::Min(50, $item.Subject.Length)) } else { "No Subject" }
                $date = if ($item.ReceivedTime) { $item.ReceivedTime.ToString("yyyy-MM-dd") } else { "No Date" }
                $className = Get-ItemTypeDescription -Class $item.Class
                Write-Log -Message "[$AccountName]   - $date ($className): $subject" -LogFile $LogFile
            }
            catch {
                Write-Log -Message "[$AccountName]   - Error reading item details" -LogFile $LogFile
            }
        }
    }
    
    Write-Log -Message "[$AccountName] === END ITEM ANALYSIS ===" -LogFile $LogFile
}

# Helper function to write summary report
function Write-SummaryReport {
    param([string]$LogFile)
    
    Write-Log -Message "=== SUMMARY REPORT ===" -LogFile $LogFile
    Write-Log -Message "Total items processed: $global:TotalItemsProcessed" -LogFile $LogFile
    Write-Log -Message "Total items moved: $global:TotalItemsMoved" -LogFile $LogFile
    Write-Log -Message "Total items skipped (rules): $global:TotalItemsSkipped" -LogFile $LogFile
    Write-Log -Message "Total items too recent: $global:TotalItemsTooRecent" -LogFile $LogFile
    Write-Log -Message "Total items with errors: $global:TotalItemsError" -LogFile $LogFile
    
    Write-Log -Message "=== ITEM TYPE BREAKDOWN ===" -LogFile $LogFile
    foreach ($type in $global:ItemTypeBreakdown.Keys | Sort-Object) {
        Write-Log -Message "$type`: $($global:ItemTypeBreakdown[$type]) items" -LogFile $LogFile
    }
    
    # Calculate leftover items
    $leftoverItems = $global:TotalItemsProcessed - $global:TotalItemsMoved
    Write-Log -Message "=== LEFTOVER ITEMS ANALYSIS ===" -LogFile $LogFile
    Write-Log -Message "Leftover items (not moved): $leftoverItems" -LogFile $LogFile
    Write-Log -Message "Breakdown of leftover items:" -LogFile $LogFile
    Write-Log -Message "  - Skipped by rules: $global:TotalItemsSkipped" -LogFile $LogFile
    Write-Log -Message "  - Too recent (< $RetentionDays days): $global:TotalItemsTooRecent" -LogFile $LogFile
    Write-Log -Message "  - Processing errors: $global:TotalItemsError" -LogFile $LogFile
    
    if ($leftoverItems -gt 0) {
        Write-Log -Message "=== RECOMMENDATIONS ===" -LogFile $LogFile
        if ($global:TotalItemsTooRecent -gt 0) {
            Write-Log -Message "- Consider reducing RetentionDays if you want to archive more recent items" -LogFile $LogFile
        }
        if ($global:TotalItemsSkipped -gt 0) {
            Write-Log -Message "- Review SkipRules in config.json if items are being skipped unexpectedly" -LogFile $LogFile
        }
        if ($global:TotalItemsError -gt 0) {
            Write-Log -Message "- Check for permission issues or corrupted items" -LogFile $LogFile
        }
    }
}

# Helper function to locate archive folders/labels for a given account
# This function uses stored paths from config.json for efficiency
# Falls back to searching if stored paths are not available or valid
function Get-ArchiveFolder {
    param($account)    # Outlook account object

    $archive = $null

    # Check if we have a stored path for this account from config.json
    # This avoids re-scanning for folders on every run
    Write-Host "  Checking for stored archive folder path for: $($account.Name)" -ForegroundColor Gray
    if ($config.ArchiveFolders -and $config.ArchiveFolders.PSObject.Properties.Name -contains $account.Name) {
        $storedPath = $config.ArchiveFolders.$($account.Name)
        Write-Host "  Found stored path: $storedPath" -ForegroundColor Green
        
        if ($storedPath -like "GmailLabel:*") {
            # Gmail label path (format: "GmailLabel:LabelName")
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
            # Root-level archive folder (format: "Root:FolderName")
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
            # Inbox-level archive folder (format: "Inbox:FolderName")
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
    # This ensures the script works even if config.json is missing or corrupted
    if (-not $archive) {
        Write-Host "  No stored path found, falling back to search..." -ForegroundColor Yellow
        Write-Host "  Searching for archive folders..." -ForegroundColor Yellow
        
        # Search for Inbox\Archive folder (common location)
        try {
            $inbox = $account.Folders.Item("Inbox")
            if ($inbox -and ($inbox.Folders | Where-Object { $_.Name -eq "Archive" })) {
                $archive = $inbox.Folders.Item("Archive")
                Write-Host "  Found Inbox\Archive folder" -ForegroundColor Green
            }
        }
        catch {}

        # Search for root-level Archive folder (alternative location)
        if (-not $archive) {
            if ($account.Folders | Where-Object { $_.Name -eq "Archive" }) {
                try { 
                    $archive = $account.Folders.Item("Archive") 
                    Write-Host "  Found root Archive folder" -ForegroundColor Green
                }
                catch {}
            }
        }

        # Search for Gmail custom label (for Gmail accounts)
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
                    
                    # Try to enumerate all folders as fallback
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

# =============================================================================
# MAIN ARCHIVING PROCESS - ACCOUNT ITERATION
# =============================================================================
# Process each Outlook account to archive old emails
# This is the core archiving logic that moves emails based on retention settings
foreach ($account in $namespace.Folders) {
    try {
        Write-Host "Processing account: $($account.Name)" -ForegroundColor Cyan
        
        # Skip non-email account types that can't be archived
        # These account types don't contain emails and would cause errors
        $skipAccountTypes = @("Internet Calendars", "SharePoint Lists", "Public Folders", "Calendar", "Contacts", "Tasks", "Notes")
        if ($skipAccountTypes -contains $account.Name) {
            $logMessage = "[$($account.Name)] Skipping non-email account type."
            Write-Log -Message $logMessage -LogFile $LogFile
            continue
        }
        
        # Get the archive folder/label for this account
        $archiveRoot = Get-ArchiveFolder $account
        if (-not $archiveRoot) {
            $logMessage = "[$($account.Name)] No 'Archive' folder found, skipping."
            Write-Log -Message $logMessage -LogFile $LogFile
            continue
        }

        # Create year/month folder structure for organized archiving
        # This keeps archived emails organized by date
        # Note: We'll create folders dynamically based on each email's received date
        # rather than creating a single folder for the current date

        # Safely retrieve the Inbox folder for this account
        $inbox = $null
        try { $inbox = $account.Folders.Item("Inbox") } catch {}
        if (-not $inbox) {
            $logMessage = "[$($account.Name)] No Inbox folder, skipping message scan."
            Write-Log -Message $logMessage -LogFile $LogFile
            continue
        }

        # Get all email items from the Inbox (Class 43 = MailItem, Class 26 = MeetingItem, Class 53 = CalendarItem)
        # Convert to static array to avoid COM object issues during iteration
        $rawItems = @()
        try {
            # Include regular emails (43), meeting invitations (26), and calendar items (53)
            $rawItems = @($inbox.Items | Where-Object { $_.Class -eq 43 -or $_.Class -eq 26 -or $_.Class -eq 53 })
            Write-Host "  Found $($rawItems.Count) items (emails + meetings + calendar items)" -ForegroundColor Cyan
        }
        catch {
            $logMessage = "[$($account.Name)] Could not retrieve mail items: $_"
            Write-Log -Message $logMessage -LogFile $LogFile
            Write-Host "  [!]  Error accessing mail items for $($account.Name): $_" -ForegroundColor Yellow
            continue
        }

        if ($rawItems.Count -eq 0) {
            $logMessage = "[$($account.Name)] No messages found to process."
            Write-Log -Message $logMessage -LogFile $LogFile
            continue
        }

        # Deduplicate emails by Subject+DateTime composite key, then sort by received time
        # This prevents processing duplicate emails and ensures chronological order
        Write-Host "  Deduplicating and sorting $($rawItems.Count) emails..." -ForegroundColor Cyan
        $seenKeys = @{}
        $deduped = foreach ($mail in $rawItems) {
            try {
                $key = "$($mail.Subject)|$($mail.ReceivedTime.ToString('o'))"
                if (-not $seenKeys.ContainsKey($key)) {
                    $seenKeys[$key] = $true
                    $mail
                }
            }
            catch {
                Write-Host "  [!]  Skipping invalid mail during deduplication" -ForegroundColor Yellow
            }
        }
        $sortedItems = $deduped | Sort-Object ReceivedTime
        Write-Host "  Processing $($sortedItems.Count) unique emails..." -ForegroundColor Green

        # Enhanced item analysis and logging
        Write-ItemAnalysis -AccountName $account.Name -Items $sortedItems -LogFile $LogFile
        
        # Process each email for archiving with enhanced tracking
        $emailCount = 0
        foreach ($mail in $sortedItems) {
            $emailCount++
            $global:TotalItemsProcessed++
            
            # Add error handling for COM object issues
            try {
                # Test if the mail object is still valid
                $null = $mail.Subject
            }
            catch {
                Write-Host "  [!]  Skipping invalid mail object at position $emailCount" -ForegroundColor Yellow
                $global:TotalItemsError++
                continue
            }
            
            # Limit to 100 emails per mailbox during dry-run for faster testing
            # This speeds up the testing process when users have many emails
            if ($DryRun -and $emailCount -gt 100) {
                $limitMessage = "[$($account.Name)] Reached 100 email limit for testing (dry-run mode)"
                Write-Log -Message $limitMessage -LogFile $LogFile
                Write-Host "  Reached 100 email limit for testing (dry-run mode)" -ForegroundColor Yellow
                break
            }
            
            # Show progress every 10 emails
            if ($emailCount % 10 -eq 0) {
                Write-Host "  Processed $emailCount emails..." -ForegroundColor Gray
            }

            # Apply skip rules from config to exclude specific emails from archiving
            # This allows users to keep important emails in their Inbox
            $skipMatch = $false
            foreach ($rule in $SkipRules) {
                if ($account.Name -eq $rule.Mailbox) {
                    foreach ($subj in $rule.Subjects) {
                        if ($mail.Subject -match [regex]::Escape($subj)) {
                            $skipMessage = "[$($account.Name)] SKIP: $($mail.ReceivedTime.ToString('yyyy-MM-dd')) : $($mail.Subject)"
                            Write-Log -Message $skipMessage -LogFile $LogFile
                            $global:TotalItemsSkipped++
                            $skipMatch = $true
                            break
                        }
                    }
                }
                if ($skipMatch) { break }
            }
            if ($skipMatch) { continue }

            # Check if email is older than retention period and should be archived
            if ($mail.ReceivedTime -lt $CutOff) {
                $logEntry = "[$($account.Name)] $($mail.ReceivedTime.ToString('yyyy-MM-dd')) : $($mail.Subject)"
                if ($DryRun) {
                    # In dry-run mode, just log what would be moved
                    $dryRunMessage = "DRY-RUN: Would move -> $logEntry"
                    Write-Log -Message $dryRunMessage -LogFile $LogFile
                    $global:TotalItemsMoved++
                }
                else {
                    # In live mode, create proper folder structure based on email's received date
                    $emailYear = $mail.ReceivedTime.ToString("yyyy")
                    $emailMonth = $mail.ReceivedTime.ToString("yyyy-MM")
                    
                    # Ensure year folder exists
                    $yearFolder = $archiveRoot.Folders | Where-Object { $_.Name -eq $emailYear }
                    if (-not $yearFolder) {
                        $archiveRoot.Folders.Add($emailYear) | Out-Null
                        $yearFolder = $archiveRoot.Folders | Where-Object { $_.Name -eq $emailYear }
                    }
                    
                    # Ensure month folder exists
                    $monthFolder = $yearFolder.Folders | Where-Object { $_.Name -eq $emailMonth }
                    if (-not $monthFolder) {
                        $yearFolder.Folders.Add($emailMonth) | Out-Null
                        $monthFolder = $yearFolder.Folders | Where-Object { $_.Name -eq $emailMonth }
                    }
                    
                    # Move the email to the correct year/month folder
                    $mail.Move($monthFolder) | Out-Null
                    $movedMessage = "MOVED: $logEntry -> $emailYear\$emailMonth"
                    Write-Log -Message $movedMessage -LogFile $LogFile
                    $global:TotalItemsMoved++
                }
            }
            else {
                # Item is too recent - log this for analysis
                $ageDays = ((Get-Date) - $mail.ReceivedTime).TotalDays
                $tooRecentMessage = "[$($account.Name)] TOO RECENT: $($mail.ReceivedTime.ToString('yyyy-MM-dd')) : $($mail.Subject) (Age: $([Math]::Round($ageDays, 1)) days)"
                Write-Log -Message $tooRecentMessage -LogFile $LogFile
                $global:TotalItemsTooRecent++
            }
        }

    }
    catch {
        $errorMessage = "[$($account.Name)] Error: $_"
        Write-Log -Message $errorMessage -LogFile $LogFile
    }
}

# =============================================================================
# PROCESSING COMPLETION
# =============================================================================
# Log completion of the archiving process
$completionMessage = "=== Completed at $(Get-Date) ==="
Write-Log -Message $completionMessage -LogFile $LogFile

# Write enhanced summary report
Write-SummaryReport -LogFile $LogFile

# =============================================================================
# POST-DRY-RUN USER INTERACTION
# =============================================================================
# After completing a dry-run, offer the user the option to switch to live mode
# This provides a safe way to test before actually moving emails
if ($DryRun) {
    Write-Host ""
    Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║   DRY-RUN COMPLETED SUCCESSFULLY!                            ║" -ForegroundColor Green
    Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Green
    Write-Host ""
    Write-Host "[SUCCESS] The dry-run has finished processing your emails." -ForegroundColor White
    Write-Host "[DOC] Log file created: $LogFile" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "[STEPS] Next steps:" -ForegroundColor Yellow
    Write-Host "1. Check the log file to review what emails would be archived" -ForegroundColor White
    Write-Host "2. Verify the archive folder structure is correct" -ForegroundColor White
    Write-Host "3. Make any adjustments to config.json if needed" -ForegroundColor White
    Write-Host ""
    
    # Ask user if they want to switch to live mode
    Write-Host "[START] Would you like to switch to live mode and run the actual archiving now?" -ForegroundColor Cyan
    Write-Host "This will move the emails that were identified in the dry-run." -ForegroundColor White
    Write-Host ""
    Write-Host "[SCHEDULE] Options:" -ForegroundColor Yellow
    Write-Host "┌─────────────────────────────────────────────────────────────────┐" -ForegroundColor Gray
    Write-Host "│ 1. YES - Switch to live mode and archive emails now             │" -ForegroundColor White
    Write-Host "│    This will move the emails that were identified in dry-run    │" -ForegroundColor Gray
    Write-Host "└─────────────────────────────────────────────────────────────────┘" -ForegroundColor Gray
    Write-Host ""
    Write-Host "┌─────────────────────────────────────────────────────────────────┐" -ForegroundColor Gray
    Write-Host "│ 2. NO - Exit and run manually later                             │" -ForegroundColor White
    Write-Host "│    You can run the script again when you're ready               │" -ForegroundColor Gray
    Write-Host "└─────────────────────────────────────────────────────────────────┘" -ForegroundColor Gray
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
            Write-Host "[OK] Configuration updated to live mode" -ForegroundColor Green
            
            Write-Host ""
            Write-Host "[!]  WARNING: This will now move emails to the archive folders!" -ForegroundColor Red
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

                        # Note: We'll create folders dynamically based on each email's received date
                        # rather than creating a single folder for the current date

                        # Safely retrieve the Inbox folder for this account
                        $inbox = $null
                        try { $inbox = $account.Folders.Item("Inbox") } catch {}
                        if (-not $inbox) {
                            $logMessage = "[$($account.Name)] No Inbox folder, skipping message scan."
                            Write-Log -Message $logMessage -LogFile $liveLogFile
                            continue
                        }

                        # Get static array of MailItems and MeetingItems
                        $rawItems = @()
                        try {
                            # Include both regular emails (43) and meeting invitations (26)
                            $rawItems = @($inbox.Items | Where-Object { $_.Class -eq 43 -or $_.Class -eq 26 -or $_.Class -eq 53 })
                            Write-Host "  Found $($rawItems.Count) items (emails + meetings + calendar items)" -ForegroundColor Cyan
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
                        Write-Host "  Deduplicating and sorting $($rawItems.Count) emails..." -ForegroundColor Cyan
                        $seenKeys = @{}
                        $deduped = foreach ($mail in $rawItems) {
                            try {
                                $key = "$($mail.Subject)|$($mail.ReceivedTime.ToString('o'))"
                                if (-not $seenKeys.ContainsKey($key)) {
                                    $seenKeys[$key] = $true
                                    $mail
                                }
                            }
                            catch {
                                Write-Host "  [!]  Skipping invalid mail during deduplication" -ForegroundColor Yellow
                            }
                        }
                        $sortedItems = $deduped | Sort-Object ReceivedTime
                        Write-Host "  Processing $($sortedItems.Count) unique emails..." -ForegroundColor Green

                        # Enhanced item analysis for live mode
                        Write-ItemAnalysis -AccountName $account.Name -Items $sortedItems -LogFile $liveLogFile
                        
                        # Process each email for archiving with enhanced tracking
                        $emailCount = 0
                        foreach ($mail in $sortedItems) {
                            $emailCount++
                            $liveEmailsProcessed++
                            
                            # Add error handling for COM object issues
                            try {
                                # Test if the mail object is still valid
                                $null = $mail.Subject
                            }
                            catch {
                                Write-Host "  [!]  Skipping invalid mail object at position $emailCount" -ForegroundColor Yellow
                                continue
                            }
                            
                            # Show progress every 10 emails
                            if ($emailCount % 10 -eq 0) {
                                Write-Host "  Processed $emailCount emails..." -ForegroundColor Gray
                            }

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
                                    # Create proper folder structure based on email's received date
                                    $emailYear = $mail.ReceivedTime.ToString("yyyy")
                                    $emailMonth = $mail.ReceivedTime.ToString("yyyy-MM")
                                    
                                    # Ensure year folder exists
                                    $yearFolder = $archiveRoot.Folders | Where-Object { $_.Name -eq $emailYear }
                                    if (-not $yearFolder) {
                                        $archiveRoot.Folders.Add($emailYear) | Out-Null
                                        $yearFolder = $archiveRoot.Folders | Where-Object { $_.Name -eq $emailYear }
                                    }
                                    
                                    # Ensure month folder exists
                                    $monthFolder = $yearFolder.Folders | Where-Object { $_.Name -eq $emailMonth }
                                    if (-not $monthFolder) {
                                        $yearFolder.Folders.Add($emailMonth) | Out-Null
                                        $monthFolder = $yearFolder.Folders | Where-Object { $_.Name -eq $emailMonth }
                                    }
                                    
                                    # Move the email to the correct year/month folder
                                    $mail.Move($monthFolder) | Out-Null
                                    $movedMessage = "MOVED: [$($account.Name)] $($mail.ReceivedTime.ToString('yyyy-MM-dd')) : $($mail.Subject) -> $emailYear\$emailMonth"
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
                                    Write-Host "  [ERROR] Error moving email: $_" -ForegroundColor Red
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
                
                # Write enhanced summary report for live mode
                Write-SummaryReport -LogFile $liveLogFile
                
                Write-Host ""
                Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Green
                Write-Host "║   LIVE ARCHIVE COMPLETED SUCCESSFULLY!                       ║" -ForegroundColor Green
                Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Green
                Write-Host ""
                Write-Host "[STATS] Results Summary:" -ForegroundColor Yellow
                Write-Host "   [EMAIL] Total emails processed: $liveEmailsProcessed" -ForegroundColor White
                Write-Host "   [OK] Total emails moved: $liveEmailsMoved" -ForegroundColor Green
                Write-Host "   [DOC] Live mode log file: $liveLogFile" -ForegroundColor Cyan
                Write-Host ""
                Write-Host "[SUCCESS] Your emails have been successfully archived!" -ForegroundColor Green
                Write-Host "[FOLDER] You can find them in your Outlook archive folders organized by year/month." -ForegroundColor White
                
            }
            else {
                Write-Host ""
                Write-Host "Live archiving cancelled by user." -ForegroundColor Yellow
                Write-Host "You can run the script again later when you're ready." -ForegroundColor White
            }
            
        }
        catch {
            Write-Host "[ERROR] Error updating configuration: $_" -ForegroundColor Red
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



# =============================================================================
# MAIN PROCESSING LOGIC (ENHANCED)
# =============================================================================
