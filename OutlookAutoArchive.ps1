<#
.SYNOPSIS
  Auto-archive Outlook emails with options from config.json
#>

Add-Type -AssemblyName Microsoft.Office.Interop.Outlook
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

# === Load config ===
$configPath = Join-Path $PSScriptRoot 'config.json'
if (-not (Test-Path $configPath)) {
    Write-Error "Config file not found: $configPath"
    exit 1
}
$config = Get-Content $configPath -Raw | ConvertFrom-Json

# === Apply config settings ===
$RetentionDays = [int]$config.RetentionDays
$DryRun = [bool]$config.DryRun
$LogPath = (Resolve-Path ($config.LogPath -replace '%USERPROFILE%', $env:USERPROFILE))
$Today = Get-Date
$CutOff = $Today.AddDays(-$RetentionDays)
$GmailLabel = $config.GmailLabel
$SkipRules = $config.SkipRules

# === Setup logging ===
if (-not (Test-Path $LogPath)) { New-Item -Path $LogPath -ItemType Directory | Out-Null }
$LogFile = Join-Path $LogPath ("ArchiveLog_" + $Today.ToString("yyyy-MM-dd_HH-mm-ss") + ".txt")

"=== Outlook Auto-Archive Dry-Run ===" | Tee-Object -FilePath $LogFile
"Retention: $RetentionDays days"       | Tee-Object -FilePath $LogFile -Append
"Cutoff: $CutOff"                       | Tee-Object -FilePath $LogFile -Append

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
            "[$($store.Name)] No 'Archive' folder found, skipping." |
            Tee-Object -FilePath $LogFile -Append
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
            "[$($store.Name)] No Inbox folder, skipping message scan." |
            Tee-Object -FilePath $LogFile -Append
            continue
        }

        # Get static array of MailItems
        $rawItems = @()
        try {
            $rawItems = @($inbox.Items | Where-Object { $_.Class -eq 43 })
        }
        catch {
            "[$($store.Name)] Could not retrieve mail items: $_" |
            Tee-Object -FilePath $LogFile -Append
            continue
        }

        if ($rawItems.Count -eq 0) {
            "[$($store.Name)] No messages found to process." |
            Tee-Object -FilePath $LogFile -Append
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
                            "[$($store.Name)] SKIP: $($mail.ReceivedTime.ToString('yyyy-MM-dd')) : $($mail.Subject)" |
                            Tee-Object -FilePath $LogFile -Append
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
                    "DRY-RUN: Would move -> $logEntry" |
                    Tee-Object -FilePath $LogFile -Append
                }
                else {
                    $mail.Move($monthFolder) | Out-Null
                    "MOVED: $logEntry" |
                    Tee-Object -FilePath $LogFile -Append
                }
            }
        }

    }
    catch {
        "[$($store.Name)] Error: $_" | Tee-Object -FilePath $LogFile -Append
    }
}

"=== Completed at $(Get-Date) ===" | Tee-Object -FilePath $LogFile -Append
