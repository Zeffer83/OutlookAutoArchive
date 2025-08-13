<#
.SYNOPSIS
  Auto-archive Outlook emails with options from config.json
#>

# Version: 1.4.0
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
    param($store)

    $archive = $null

    # Inbox\Archive
    try {
        $inbox = $store.Folders.Item("Inbox")
        if ($inbox -and ($inbox.Folders | Where-Object { $_.Name -eq "Archive" })) {
            $archive = $inbox.Folders.Item("Archive")
        }
    }
    catch {}

    # Root-level Archive
    if (-not $archive) {
        if ($store.Folders | Where-Object { $_.Name -eq "Archive" }) {
            try { $archive = $store.Folders.Item("Archive") } catch {}
        }
    }

    # Gmail custom label
    if (-not $archive -and $GmailLabel) {
        foreach ($folder in $store.Folders) {
            if ($folder.Name -eq $GmailLabel) { $archive = $folder; break }
        }
    }

    return $archive
}

foreach ($store in $namespace.Folders) {
    try {
        $archiveRoot = Get-ArchiveFolder $store
        if (-not $archiveRoot) {
            $logMessage = "[$($store.Name)] No 'Archive' folder found, skipping."
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
        try { $inbox = $store.Folders.Item("Inbox") } catch {}
        if (-not $inbox) {
            $logMessage = "[$($store.Name)] No Inbox folder, skipping message scan."
            Write-Log -Message $logMessage -LogFile $LogFile
            continue
        }

        # Get static array of MailItems
        $rawItems = @()
        try {
            $rawItems = @($inbox.Items | Where-Object { $_.Class -eq 43 })
        }
        catch {
            $logMessage = "[$($store.Name)] Could not retrieve mail items: $_"
            Write-Log -Message $logMessage -LogFile $LogFile
            continue
        }

        if ($rawItems.Count -eq 0) {
            $logMessage = "[$($store.Name)] No messages found to process."
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
                if ($store.Name -eq $rule.Mailbox) {
                    foreach ($subj in $rule.Subjects) {
                        if ($mail.Subject -match [regex]::Escape($subj)) {
                            $skipMessage = "[$($store.Name)] SKIP: $($mail.ReceivedTime.ToString('yyyy-MM-dd')) : $($mail.Subject)"
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
                $logEntry = "[$($store.Name)] $($mail.ReceivedTime.ToString('yyyy-MM-dd')) : $($mail.Subject)"
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
        $errorMessage = "[$($store.Name)] Error: $_"
        Write-Log -Message $errorMessage -LogFile $LogFile
    }
}

$completionMessage = "=== Completed at $(Get-Date) ==="
Write-Log -Message $completionMessage -LogFile $LogFile
