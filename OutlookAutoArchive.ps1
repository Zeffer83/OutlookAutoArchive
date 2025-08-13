<#
.SYNOPSIS
  Moves emails older than X days from Inbox to Archive\<yyyy>\<yyyy-MM>.
  Dry-run mode enabled — flip $DryRun to $false for live moves.
#>

Add-Type -AssemblyName Microsoft.Office.Interop.Outlook
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

# === CONFIG ===
$RetentionDays = 14
$DryRun = $true   # change to $false for live mode
$LogPath = "$env:USERPROFILE\Documents\OutlookAutoArchiveLogs"
$Today = Get-Date
$CutOff = $Today.AddDays(-$RetentionDays)

# === Setup logging ===
if (-not (Test-Path $LogPath)) { New-Item -Path $LogPath -ItemType Directory | Out-Null }
$LogFile = Join-Path $LogPath ("ArchiveLog_" + $Today.ToString("yyyy-MM-dd_HH-mm-ss") + ".txt")

"=== Outlook Auto-Archive Dry-Run ===" | Tee-Object -FilePath $LogFile
"Retention: $RetentionDays days"       | Tee-Object -FilePath $LogFile -Append
"Cutoff: $CutOff"                       | Tee-Object -FilePath $LogFile -Append

function Get-ArchiveFolder {
    param($store)

    $archive = $null

    # 1. Try Inbox\Archive
    try {
        $inbox = $store.Folders.Item("Inbox")
        if ($inbox -and ($inbox.Folders | Where-Object { $_.Name -eq "Archive" })) {
            $archive = $inbox.Folders.Item("Archive")
        }
    }
    catch {}

    # 2. Try root-level Archive
    if (-not $archive) {
        if ($store.Folders | Where-Object { $_.Name -eq "Archive" }) {
            try { $archive = $store.Folders.Item("Archive") } catch {}
        }
    }

    # 3. Gmail custom: look for "OutlookArchive" label
    if (-not $archive) {
        foreach ($folder in $store.Folders) {
            if ($folder.Name -eq "OutlookArchive") { $archive = $folder; break }
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

            # --- SKIP LOGIC FOR "Surinder (Shared)" ---
            if ($store.Name -eq "Surinder (Shared)") {
                if ($mail.Subject -match "Medite\s+offline\s+monitoring" -or
                    $mail.Subject -match "Medite\s+Offline\s+Monitoring\s+Service") {
                    "[$($store.Name)] SKIP: $($mail.ReceivedTime.ToString('yyyy-MM-dd')) : $($mail.Subject)" |
                    Tee-Object -FilePath $LogFile -Append
                    continue
                }
            }
            # ------------------------------------------

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
