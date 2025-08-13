# Outlook Auto Archive Script

A PowerShell script that automatically archives emails older than a specified number of days from your Outlook Inbox to organized year/month folders.

**Author**: Ryan Zeffiretti  
**Version**: 1.1.0  
**License**: MIT

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
3. Copy `config.example.json` to `config.json` and customize the settings
4. The script is ready to run - no additional installation required

## Configuration

The script uses a `config.json` file for configuration. Edit this file to customize the script behavior:

```json
{
  "RetentionDays": 14,
  "DryRun": true,
  "LogPath": "%USERPROFILE%\\Documents\\OutlookAutoArchiveLogs",
  "GmailLabel": "OutlookArchive",
  "SkipRules": [
    {
      "Mailbox": "MailBox Name",
      "Subjects": ["Message Subject", "Message Subject"]
    }
  ]
}
```

### Configuration Options

- **`RetentionDays`**: Number of days to keep emails in Inbox before archiving
- **`DryRun`**: When `true`, shows what would be moved without actually moving emails
- **`LogPath`**: Directory where log files are stored (supports `%USERPROFILE%` variable)
- **`GmailLabel`**: Custom Gmail label name for archive folder (optional)
- **`SkipRules`**: Array of rules to skip specific emails by mailbox and subject patterns

## Usage

### Basic Usage

```powershell
# Run in dry-run mode (recommended first time)
# Edit config.json to set "DryRun": true
.\OutlookAutoArchive.ps1

# Run in live mode (actually moves emails)
# Edit config.json to set "DryRun": false
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

To skip specific emails, add rules to the `config.json` file:

```json
"SkipRules": [
    {
        "Mailbox": "Your Mailbox Name",
            "Subjects": [
                "Subject Pattern 1",
                "Subject Pattern 2"
            ]
    }
]
```

The script will automatically skip emails that match the specified mailbox and subject patterns.

### Custom Archive Locations

The script automatically detects archive folders in multiple locations:

- `Inbox\Archive`
- Root-level `Archive`
- Custom Gmail labels (configured via `GmailLabel` in config.json)

To add support for additional locations, modify the `Get-ArchiveFolder` function in the script.

### Setting Up Gmail Labels

If you're using Gmail with Outlook, you can create custom labels for archiving:

#### In Gmail:
1. **Create a Label**:
   - Open Gmail in your web browser
   - Click the gear icon (Settings) → "See all settings"
   - Go to the "Labels" tab
   - Click "Create new label"
   - Name it (e.g., "OutlookArchive" or your preferred name)
   - Click "Create"

2. **Label Structure** (Optional):
   - You can create nested labels like "OutlookArchive/2024/2024-12"
   - The script will automatically create year/month sub-labels

#### In Outlook:
1. **Sync the Label**:
   - The Gmail label should automatically appear in Outlook
   - It may take a few minutes to sync
   - Look for the label in your folder list

2. **Configure the Script**:
   - Set `GmailLabel` in your `config.json` to match your Gmail label name
   - Example: `"GmailLabel": "OutlookArchive"`

#### Troubleshooting Gmail Labels:
- **Label not appearing**: Try refreshing Outlook or restarting it
- **Sync issues**: Check your Gmail IMAP settings in Outlook
- **Permission errors**: Ensure you have full access to your Gmail account

## Troubleshooting

### Common Issues

1. **"Access Denied" errors**: Ensure Outlook is running and you have permissions
2. **No Archive folder found**: Create an Archive folder in your Inbox or root level
3. **Script won't run**: Check PowerShell execution policy: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned`
4. **Config file not found**: Ensure `config.json` exists in the same directory as the script
5. **Invalid JSON**: Check that your `config.json` file has valid JSON syntax

### Debug Mode

Enable verbose logging by adding `-Verbose` to PowerShell commands or modifying the script to include more detailed output.

## Contributing

This is an "as-is" script created for personal use. While contributions are welcome, please note that this project is not actively maintained beyond personal needs. If you find issues or want to add features, feel free to fork and modify for your own use.

If you do find critical bugs, you can:

1. Fork the repository
2. Fix the issue
3. Submit a pull request (though response may be limited)

## License

This project is open source and available under the [MIT License](LICENSE).

**Author**: Ryan Zeffiretti (rzeffiretti@gmail.com)

## ⚠️ Important Disclaimers

### 🔒 Data Safety Warning

**ALWAYS BACKUP YOUR DATA BEFORE USE!** While this script includes safety features like dry-run mode and comprehensive logging, it's your responsibility to ensure you have proper backups of your email data before using this tool.

### 🛡️ No Warranty

This software is provided "AS IS" without warranty of any kind. The author makes no representations or warranties about the accuracy, reliability, completeness, or suitability of this software for any purpose.

### 🛡️ Limitation of Liability

The author shall not be liable for any direct, indirect, incidental, special, consequential, or punitive damages, including but not limited to:

- Loss of data or emails
- Email corruption or deletion
- System corruption
- Any other damages arising from the use of this software

### 📋 User Responsibility

By using this software, you acknowledge that:

- You have backed up your email data before use
- You understand the risks involved in email operations
- You accept full responsibility for any consequences
- You will test the software in dry-run mode first
- You have read and understood all disclaimers

### ⚠️ Disclaimer

This script modifies your Outlook email structure. Always test in dry-run mode first and ensure you have backups of important emails before running in live mode.

**Note**: This is an "as-is" script created for personal use. Use at your own risk and test thoroughly in your environment.

## Support

If you encounter issues:

1. Check the log files for detailed error messages
2. Ensure Outlook is properly configured
3. Verify you have necessary permissions
4. Test with dry-run mode enabled

---

**Note**: This script is designed for personal use and should be tested thoroughly in your environment before production use. This is version 1.1.0 and is provided "as-is" with no planned updates unless critical issues are found.
