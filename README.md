# Outlook Auto Archive Script

A PowerShell script that automatically archives emails older than a specified number of days from your Outlook Inbox to organized year/month folders.

## Features

- **Automatic Email Archiving**: Moves emails older than X days from Inbox to Archive folders
- **Organized Structure**: Creates year/month folder hierarchy (e.g., `Archive\2024\2024-12`)
- **Dry-Run Mode**: Test the script safely without actually moving emails
- **Comprehensive Logging**: Detailed logs of all operations with timestamps
- **Multi-Account Support**: Works with multiple Outlook accounts/stores
- **Smart Folder Detection**: Automatically finds Archive folders in various locations
- **Duplicate Prevention**: Handles duplicate emails intelligently
- **Custom Skip Rules**: Built-in logic to skip specific emails (e.g., monitoring alerts)

## Prerequisites

- Windows with PowerShell 5.1 or later
- Microsoft Outlook installed and configured
- Outlook COM Interop permissions

## Installation

1. Clone or download this repository
2. Ensure Outlook is installed and configured on your system
3. The script is ready to run - no additional installation required

## Configuration

Edit the configuration section in `OutlookAutoArchive.ps1`:

```powershell
# === CONFIG ===
$RetentionDays = 14        # Keep emails in Inbox for 14 days
$DryRun = $true           # Set to $false for live mode
$LogPath = "$env:USERPROFILE\Documents\OutlookAutoArchiveLogs"
```

### Configuration Options

- **`$RetentionDays`**: Number of days to keep emails in Inbox before archiving
- **`$DryRun`**: When `$true`, shows what would be moved without actually moving emails
- **`$LogPath`**: Directory where log files are stored

## Usage

### Basic Usage

```powershell
# Run in dry-run mode (recommended first time)
.\OutlookAutoArchive.ps1

# Run in live mode (actually moves emails)
# First edit the script to set $DryRun = $false
.\OutlookAutoArchive.ps1
```

### Scheduled Execution

To run automatically, create a Windows Task Scheduler task:

1. Open Task Scheduler
2. Create Basic Task
3. Set trigger (e.g., daily at 2 AM)
4. Action: Start a program
5. Program: `powershell.exe`
6. Arguments: `-ExecutionPolicy Bypass -File "C:\path\to\OutlookAutoArchive.ps1"`

## How It Works

1. **Connects to Outlook**: Uses COM Interop to access Outlook
2. **Finds Archive Folder**: Searches for Archive folder in multiple locations:
   - `Inbox\Archive`
   - Root-level `Archive`
   - Gmail-style `OutlookArchive`
3. **Creates Folder Structure**: Automatically creates year/month folders
4. **Scans Inbox**: Processes all emails in the Inbox
5. **Moves Old Emails**: Moves emails older than retention period to appropriate archive folder
6. **Logs Everything**: Records all operations to timestamped log files

## Archive Structure

The script creates an organized folder structure:

```
Archive/
├── 2024/
│   ├── 2024-01/
│   ├── 2024-02/
│   └── ...
├── 2023/
│   ├── 2023-12/
│   └── ...
```

## Logging

Logs are stored in: `%USERPROFILE%\Documents\OutlookAutoArchiveLogs\`

Each run creates a timestamped log file: `ArchiveLog_YYYY-MM-DD_HH-mm-ss.txt`

Log entries include:
- Configuration settings
- Folder creation operations
- Email movement details
- Errors and warnings
- Completion timestamp

## Safety Features

- **Dry-Run Mode**: Test without making changes
- **Duplicate Prevention**: Avoids moving duplicate emails
- **Error Handling**: Graceful handling of missing folders or permissions
- **Backup Logging**: All operations are logged before execution

## Customization

### Adding Skip Rules

To skip specific emails, add conditions in the main processing loop:

```powershell
# Example: Skip emails with specific subjects
if ($mail.Subject -match "Your Pattern Here") {
    "SKIP: $($mail.Subject)" | Tee-Object -FilePath $LogFile -Append
    continue
}
```

### Custom Archive Locations

Modify the `Get-ArchiveFolder` function to support additional archive folder locations.

## Troubleshooting

### Common Issues

1. **"Access Denied" errors**: Ensure Outlook is running and you have permissions
2. **No Archive folder found**: Create an Archive folder in your Inbox or root level
3. **Script won't run**: Check PowerShell execution policy: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned`

### Debug Mode

Enable verbose logging by adding `-Verbose` to PowerShell commands or modifying the script to include more detailed output.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

This project is open source and available under the [MIT License](LICENSE).

## Disclaimer

This script modifies your Outlook email structure. Always test in dry-run mode first and ensure you have backups of important emails before running in live mode.

## Support

If you encounter issues:
1. Check the log files for detailed error messages
2. Ensure Outlook is properly configured
3. Verify you have necessary permissions
4. Test with dry-run mode enabled

---

**Note**: This script is designed for personal use and should be tested thoroughly in your environment before production use.
