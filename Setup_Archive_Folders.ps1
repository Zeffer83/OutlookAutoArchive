<#
.SYNOPSIS
    Setup script to automatically create archive folders and labels for Outlook Auto Archive
.DESCRIPTION
    This script automatically creates the necessary archive folders and labels for all email accounts
    in Outlook, making setup much easier for users. It supports regular email accounts and Gmail.
.PARAMETER GmailLabel
    The name of the Gmail label to create (default: "OutlookArchive")
.PARAMETER CreateInInbox
    Create archive folder inside Inbox instead of root level
.PARAMETER DryRun
    Show what would be created without actually creating folders
.EXAMPLE
    .\Setup_Archive_Folders.ps1
    Creates archive folders for all accounts using default settings
.EXAMPLE
    .\Setup_Archive_Folders.ps1 -GmailLabel "MyArchive" -CreateInInbox
    Creates "MyArchive" Gmail label and Inbox\Archive folders for other accounts
#>

param(
    [string]$GmailLabel = "OutlookArchive",
    [switch]$CreateInInbox,
    [switch]$DryRun
)

# Version: 1.0.0
# Author: Ryan Zeffiretti
# Description: Setup script for Outlook Auto Archive folders and labels

Write-Host "=== Outlook Auto Archive - Folder Setup Script ===" -ForegroundColor Cyan
Write-Host "Version: 1.0.0" -ForegroundColor Gray
Write-Host "Author: Ryan Zeffiretti" -ForegroundColor Gray
Write-Host ""

if ($DryRun) {
    Write-Host "DRY-RUN MODE: No folders will be created" -ForegroundColor Yellow
    Write-Host ""
}

# Check if Outlook is running
try {
    $outlookProcesses = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
    if (-not $outlookProcesses) {
        Write-Host "❌ Outlook is not running. Please start Outlook and try again." -ForegroundColor Red
        Write-Host "The script requires Outlook to be running to access email data." -ForegroundColor Yellow
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
Write-Host "Scanning email accounts..." -ForegroundColor Cyan

$accountsProcessed = 0
$foldersCreated = 0
$errors = 0

foreach ($account in $namespace.Folders) {
    try {
        Write-Host ""
        Write-Host "Processing account: $($account.Name)" -ForegroundColor Cyan
        
        $accountProcessed = $false
        $accountFoldersCreated = 0
        
        # Check if this looks like a Gmail account
        $isGmail = $account.Name -like "*@gmail.com" -or $account.Name -like "*@googlemail.com"
        
        if ($isGmail) {
            Write-Host "  Detected Gmail account" -ForegroundColor Gray
            
            # Check if Gmail label already exists
            $existingLabel = $null
            try {
                $existingLabel = $account.Folders.Item($GmailLabel)
            }
            catch {}
            
            if ($existingLabel) {
                Write-Host "  ✅ Gmail label '$GmailLabel' already exists" -ForegroundColor Green
                $accountProcessed = $true
            }
            else {
                if ($DryRun) {
                    Write-Host "  🔍 DRY-RUN: Would create Gmail label '$GmailLabel'" -ForegroundColor Yellow
                    $accountFoldersCreated++
                }
                else {
                    try {
                        $newLabel = $account.Folders.Add($GmailLabel)
                        Write-Host "  ✅ Created Gmail label '$GmailLabel'" -ForegroundColor Green
                        $accountFoldersCreated++
                        $accountProcessed = $true
                    }
                    catch {
                        Write-Host "  ❌ Failed to create Gmail label: $_" -ForegroundColor Red
                        $errors++
                    }
                }
            }
        }
        else {
            Write-Host "  Detected regular email account" -ForegroundColor Gray
            
            # Try to create archive folder
            $archiveFolder = $null
            $archivePath = ""
            
            if ($CreateInInbox) {
                # Create in Inbox\Archive
                try {
                    $inbox = $account.Folders.Item("Inbox")
                    if ($inbox) {
                        # Check if Archive folder already exists in Inbox
                        try {
                            $archiveFolder = $inbox.Folders.Item("Archive")
                            $archivePath = "Inbox\Archive"
                        }
                        catch {
                            if ($DryRun) {
                                Write-Host "  🔍 DRY-RUN: Would create Inbox\Archive folder" -ForegroundColor Yellow
                                $accountFoldersCreated++
                            }
                            else {
                                $archiveFolder = $inbox.Folders.Add("Archive")
                                Write-Host "  ✅ Created Inbox\Archive folder" -ForegroundColor Green
                                $accountFoldersCreated++
                            }
                        }
                        $accountProcessed = $true
                    }
                }
                catch {
                    Write-Host "  ⚠️  Could not access Inbox folder" -ForegroundColor Yellow
                }
            }
            
            # If not created in Inbox or CreateInInbox is false, try root level
            if (-not $archiveFolder) {
                try {
                    $archiveFolder = $account.Folders.Item("Archive")
                    $archivePath = "Archive"
                    Write-Host "  ✅ Archive folder already exists at root level" -ForegroundColor Green
                    $accountProcessed = $true
                }
                catch {
                    if ($DryRun) {
                        Write-Host "  🔍 DRY-RUN: Would create root-level Archive folder" -ForegroundColor Yellow
                        $accountFoldersCreated++
                    }
                    else {
                        try {
                            $archiveFolder = $account.Folders.Add("Archive")
                            Write-Host "  ✅ Created root-level Archive folder" -ForegroundColor Green
                            $accountFoldersCreated++
                            $accountProcessed = $true
                        }
                        catch {
                            Write-Host "  ❌ Failed to create Archive folder: $_" -ForegroundColor Red
                            $errors++
                        }
                    }
                }
            }
        }
        
        if ($accountProcessed) {
            $accountsProcessed++
            $foldersCreated += $accountFoldersCreated
        }
        else {
            Write-Host "  ⚠️  Could not process this account" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "  ❌ Error processing account '$($account.Name)': $_" -ForegroundColor Red
        $errors++
    }
}

Write-Host ""
Write-Host "=== Setup Summary ===" -ForegroundColor Cyan
Write-Host "Accounts processed: $accountsProcessed" -ForegroundColor White
Write-Host "Folders/labels created: $foldersCreated" -ForegroundColor White
Write-Host "Errors encountered: $errors" -ForegroundColor $(if ($errors -gt 0) { "Red" } else { "Green" })

if ($DryRun) {
    Write-Host ""
    Write-Host "💡 To actually create the folders, run without -DryRun parameter" -ForegroundColor Yellow
}
else {
    Write-Host ""
    if ($foldersCreated -gt 0) {
        Write-Host "✅ Setup completed successfully!" -ForegroundColor Green
        Write-Host ""
        Write-Host "Next steps:" -ForegroundColor Cyan
        Write-Host "1. Run OutlookAutoArchive.exe to test the setup" -ForegroundColor White
        Write-Host "2. Check the log files to verify everything works" -ForegroundColor White
        Write-Host "3. Set up scheduled execution if desired" -ForegroundColor White
    }
    else {
        Write-Host "ℹ️  No new folders were created (they may already exist)" -ForegroundColor Blue
    }
}

Write-Host ""
Write-Host "For Gmail users:" -ForegroundColor Cyan
Write-Host "- Make sure IMAP is enabled in Gmail settings" -ForegroundColor White
Write-Host "- Check 'Show in IMAP' for your labels in Gmail web interface" -ForegroundColor White
Write-Host "- It may take a few minutes for labels to sync to Outlook" -ForegroundColor White

Write-Host ""
Write-Host "=== Setup Complete ===" -ForegroundColor Cyan
