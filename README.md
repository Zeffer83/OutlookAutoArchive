# Outlook Auto Archive Script

A PowerShell script that automatically archives emails older than a specified number of days from your Outlook Inbox to organized year/month folders. The script creates a structured archive system with folders organized by year and month (e.g., `Archive\2025\2025-08`) for easy email retrieval and management.

**Author**: Ryan Zeffiretti  
**Version**: 2.8.2  
**License**: MIT

## Features

- **Automatic Email Archiving**: Moves emails older than X days from Inbox to Archive folders
- **Organized Structure**: Creates year/month folder hierarchy (e.g., `Archive\2025\2025-08`)
- **Dry-Run Mode**: Test the script safely without actually moving emails
- **Comprehensive Logging**: Detailed logs of all operations with timestamps
- **Multi-Account Support**: Works with multiple Outlook accounts/stores
- **Smart Folder Detection**: Automatically finds Archive folders in various locations
- **Duplicate Prevention**: Handles duplicate emails intelligently
- **Custom Skip Rules**: Built-in logic to skip specific emails (e.g., monitoring alerts)
- **Outlook Status Check**: Automatically verifies Outlook is running before execution
- **Enhanced Error Handling**: Improved logging and error recovery
- **Continuous Monitoring**: Optional periodic archiving while computer is running (perfect for users who leave their computer on for days)
- **Smart Monitoring**: Starts when Outlook opens, then runs every 4-24 hours with graceful handling when Outlook is closed

## Safety Features

- **Dry-Run Mode**: Test without making changes
- **Duplicate Prevention**: Avoids moving duplicate emails
- **Error Handling**: Graceful handling of missing folders or permissions
- **Backup Logging**: All operations are logged before execution

## Prerequisites

- Windows with PowerShell 5.1 or later
- Microsoft Outlook installed and configured
- Outlook COM Interop permissions

## Windows Security

### Automatic Unblocking

The application automatically detects and attempts to unblock itself if Windows has blocked it due to being downloaded from the internet. This is a normal Windows security feature.

### Manual Unblocking

If automatic unblocking fails, you can manually unblock the executable:

1. Right-click on `OutlookAutoArchive.exe`
2. Select "Properties"
3. Check the "Unblock" checkbox at the bottom of the dialog
4. Click "OK"
5. Run the application again

This is a one-time process - once unblocked, the file will run normally on subsequent executions.

## Installation

1. **Download the files** from this repository
2. **Extract to a folder** (e.g., `C:\OutlookAutoArchive\`)
3. **Ensure Outlook is installed** and configured on your system
4. **Create the main Archive folder** (see Setup Requirements section below)
5. **Run the executable** to test the setup

**Files included:**

- `OutlookAutoArchive.exe` - **Single executable (RECOMMENDED for all users) - FULLY TESTED AND WORKING**
- `config.example.json` - Example configuration file
- `config.json` - Your configuration file (auto-created on first run)

**Source Code (for developers only):**

- `OutlookAutoArchive.ps1` - PowerShell script source code (for developers and advanced users)

**First Run Setup:**

The script will automatically create a `config.json` file on first run if one doesn't exist. It will either copy from `config.example.json` if available, or create a default configuration with safe settings (DryRun = true).

**Installation Options:**

The script offers three installation locations during first-run setup:

1. **User Documents** (`C:\Users\username\OutlookAutoArchive\`) - **Recommended**

   - User-specific installation (no admin permissions required)
   - Easy to find and manage
   - Works with all Windows user types

2. **Custom Location** - Choose your own folder

   - Full control over installation path
   - Useful for network drives or specific directories

3. **Current Location** - Keep everything where it is now
   - No file movement required
   - Good for portable installations

**Recommended Setup Process:**

1. **Double-click** `OutlookAutoArchive.exe` to start the first-run setup
2. **Choose installation location** (User Documents recommended)
3. **Follow the guided setup** to configure archive folders and scheduling
4. **Check the log files** in the `Logs` folder within your installation directory
5. **Review and edit** `config.json` if needed
6. **Test again** to ensure everything works
7. **Set up scheduled execution** (see Scheduled Execution section)

## Configuration

The script uses a `config.json` file for configuration. Edit this file to customize the script behavior:

```json
{
  "RetentionDays": 14,
  "DryRun": true,
  "LogPath": ".\\Logs",
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
- **`OnFirstRun`**: Set to `true` to enable interactive first-run setup (auto-set to `false` after setup)
- **`ArchiveFolders`**: Automatically populated with detected archive folder paths for faster access
- **`SkipRules`**: Array of rules to skip specific emails by mailbox and subject patterns

## ✅ **Status**: FULLY TESTED AND WORKING - Version 2.7.0

This script has been thoroughly tested and is working perfectly with all email account types (Gmail, Outlook, Exchange, etc.). It successfully detects archive folders, processes emails, and applies skip rules correctly.

**Test Results**: Successfully processed 1,000+ emails across multiple accounts with proper archive folder detection and skip rule functionality.

**Latest Improvements (v2.7.0)**:

- ✅ **Simplified Scheduling**: Streamlined to just 3 clear options - Daily, Startup + Monitoring, or Skip
- ✅ **Startup + Monitoring**: Combines system startup with 4-hour periodic monitoring for optimal archiving
- ✅ **Graceful Outlook Handling**: Script gracefully skips runs when Outlook is not available (no failed tasks)
- ✅ **Better User Experience**: Cleaner interface with fewer confusing options

**Previous Improvements (v2.6.0)**:

- ✅ **Smart Monitoring**: New scheduling option that starts when Outlook opens, then runs every 4-24 hours
- ✅ **Graceful Outlook Handling**: Script gracefully skips scheduled runs when Outlook is not available (no more failed tasks)
- ✅ **Enhanced Continuous Monitoring**: Improved error handling for all scheduled tasks
- ✅ **Better User Experience**: Clear distinction between different monitoring options

**Previous Improvements (v2.5.0)**:

- ✅ **Continuous Monitoring**: New scheduling option that runs every 4-24 hours while computer is on (perfect for users who leave their computer running for days)
- ✅ **Customizable Monitoring Interval**: Users can choose how often the script runs (1-24 hours)
- ✅ **Smart Scheduling**: Addresses the issue where emails wouldn't be archived if Outlook wasn't restarted for days
- ✅ **Enhanced User Experience**: Clear explanations of each scheduling option and their benefits

**Previous Improvements (v2.4.0)**:

- ✅ **Archive Folder Storage**: Archive folder paths are stored in config.json for faster subsequent runs
- ✅ **Performance Optimization**: Eliminates re-scanning of archive folders on every run
- ✅ **Enhanced User Feedback**: Shows which archive folders were discovered and stored
- ✅ **Smart Path Detection**: Automatically detects and stores Gmail labels, Inbox folders, and root folders
- ✅ **Backward Compatibility**: Falls back to searching if stored paths become invalid

**Previous Improvements (v2.2.0)**:

- ✅ **Streamlined Installation**: Only essential files copied (exe + config example)
- ✅ **User-Friendly README.txt**: Simple, clear instructions instead of complex markdown
- ✅ **Clean Installation**: No development files or repository clutter
- ✅ **Professional Appearance**: Minimal, focused installation directory
- ✅ **Auto-Skip Non-Email Accounts**: Automatically skips Internet Calendars, SharePoint Lists, etc.
- ✅ **Improved Logging**: Logs stored in installation directory (no OneDrive confusion)

**New in v2.0.0**: Ultimate single-file experience! Removed all standalone setup scripts - everything is now integrated into the main executable with a complete guided first-run setup.

## How It Works

1. **Connects to Outlook**: Uses COM Interop to access Outlook
2. **Finds Archive Folder**: Searches for Archive folder in multiple locations:
   - `Inbox\Archive`
   - Root-level `Archive`
   - Custom Gmail labels (configured via `GmailLabel` in config.json)
3. **Creates Folder Structure**: Automatically creates year/month folders (you only need to create the main Archive folder)
4. **Scans Inbox**: Processes all emails in the Inbox
5. **Moves Old Emails**: Moves emails older than retention period to appropriate archive folder
6. **Logs Everything**: Records all operations to timestamped log files

## Setup Requirements

**EASY SETUP**: Use the provided setup script to automatically create archive folders and labels!

### Option 1: Automatic Setup (RECOMMENDED)

1. **Double-click** `OutlookAutoArchive.exe`
2. **Follow the guided first-run setup** to create archive folders and labels
3. **The script will automatically**:
   - Detect all your email accounts
   - Create Gmail labels for Gmail accounts
   - Create Archive folders for regular email accounts
   - Handle all the complex setup automatically

### Option 2: Manual Setup (Advanced Users)

If you prefer to create folders manually:

#### Create Archive folder in Inbox

1. Right-click on your Inbox
2. Select "New Folder"
3. Name it "Archive"

#### Create Archive folder at root level

1. Right-click on your email account name
2. Select "New Folder"
3. Name it "Archive"

#### Use Gmail labels (see Gmail setup section below)

Configure a Gmail label in your `config.json` file.

**The main script will automatically create**:

- Year folders (e.g., "2025", "2024")
- Month folders (e.g., "2025-08", "2025-07")

## Archive Structure

The script creates an organized folder structure:

```
Archive/
├── 2025/
│   ├── 2025-08/
│   ├── 2025-07/
│   └── ...
├── 2024/
│   ├── 2024-12/
│   └── ...
```

## Usage

### Option 1: Using the Executable (Recommended)

The easiest way to run the script is using the provided executable:

**Method A: Run the executable directly**

```powershell
# Run the executable (it will use config.json automatically)
.\OutlookAutoArchive.exe
```

**Benefits of using the executable:**

- No PowerShell execution policy issues
- Double-click to run
- Works with Windows Task Scheduler
- No need to open PowerShell
- Integrated Outlook status checking
- Can generate convenience batch files on first run

### Option 2: Using PowerShell Script (Developers Only)

```powershell
# Run in dry-run mode (recommended first time)
# Edit config.json to set "DryRun": true
.\OutlookAutoArchive.ps1

# Run in live mode (actually moves emails)
# Edit config.json to set "DryRun": false
.\OutlookAutoArchive.ps1
```

**Note**: The PowerShell script is provided for developers and advanced users. Most users should use the executable version.

### Scheduled Execution

The first-run setup offers three simple scheduling options:

#### Method 1: Daily at Specific Time

Choose "Daily at a specific time" during first-run setup for traditional scheduling.

This creates a scheduled task that:

- Runs once per day at your chosen time (e.g., 2:00 AM)
- Perfect for users who want predictable, scheduled archiving
- Works well with regular computer usage patterns

#### Method 2: Startup + Monitoring (Recommended)

Choose "When Outlook starts + every 4 hours" during first-run setup for the best solution.

This creates a scheduled task that:

- Starts when the computer boots (with 30-second delay)
- Runs every 4 hours while the computer is on
- Gracefully skips runs when Outlook is not available (no failed tasks)
- Perfect for users who want archiving only when Outlook is available
- Best option for users who close Outlook occasionally

**Benefits:**

- No failed scheduled tasks when Outlook is closed
- Automatic archiving when Outlook is available
- Smart detection of Outlook availability
- Optimal for most users who want reliable archiving

#### Method 3: Manual Setup (Advanced Users)

If you prefer to set up scheduling manually:

1. **Open Task Scheduler** (search in Start menu)
2. **Click "Create Basic Task"** in the right panel
3. **Name**: `Outlook Auto Archive`
4. **Description**: `Automatically archive old emails from Outlook`
5. **Trigger**: Choose your schedule (e.g., Daily at 2:00 AM)
6. **Action**: Start a program
7. **Program/script**: `C:\path\to\OutlookAutoArchive.exe`
8. **Arguments**: (leave empty)
9. **Finish**: Review settings and click Finish

**Command Line Alternative:**

```cmd
schtasks /create /tn "Outlook Auto Archive" /tr "C:\path\to\OutlookAutoArchive.exe" /sc daily /st 02:00 /f
```

#### Testing Your Scheduled Task

1. **Run manually first**: Right-click the task → "Run"
2. **Check logs**: Verify log files are created
3. **Monitor execution**: Check task history in Task Scheduler
4. **Adjust timing**: Ensure Outlook is running when task executes

## Logging

Logs are stored in: `Logs\` folder within your installation directory

Each run creates a timestamped log file: `ArchiveLog_YYYY-MM-DD_HH-mm-ss.txt`

Log entries include:

- Configuration settings
- Folder creation operations
- Email movement details
- Errors and warnings
- Completion timestamp

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

### Archive Folder Detection

The script automatically detects archive folders in multiple locations:

- `Inbox\Archive`
- Root-level `Archive`
- Custom Gmail labels (configured via `GmailLabel` in config.json)

To add support for additional locations, modify the `Get-ArchiveFolder` function in the script.

### Setting Up Gmail Labels

If you're using Gmail with Outlook, you can create custom labels for archiving:

#### In Gmail:

1. **Enable IMAP** (Required for Outlook sync):

   - Open Gmail in your web browser
   - Click the gear icon (Settings) → "See all settings"
   - Go to the "Forwarding and POP/IMAP" tab
   - In the "IMAP access" section, select "Enable IMAP"
   - Click "Save Changes"

2. **Create a Label**:

   - Go to the "Labels" tab in Gmail settings
   - Click "Create new label"
   - Name it (e.g., "OutlookArchive" or your preferred name)
   - Click "Create"

3. **Show Label in IMAP** (Required for Outlook sync):

   - In the "Labels" tab, find your newly created label
   - Check the box under "Show in IMAP" for your label
   - This ensures the label appears in Outlook

4. **Label Structure** (Optional):
   - You can create nested labels like "OutlookArchive/2025/2025-08"
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
- **IMAP not enabled**: Make sure IMAP is enabled in Gmail settings (Forwarding and POP/IMAP tab)
- **Labels not syncing**: Wait a few minutes for Gmail to sync labels to Outlook, or restart Outlook
- **Label not showing in Outlook**: Ensure "Show in IMAP" is checked for your label in Gmail settings
- **Custom labels missing**: Only labels marked "Show in IMAP" will appear in Outlook

## Troubleshooting

### Common Issues

1. **"Access Denied" errors**: Ensure Outlook is running and you have permissions
2. **No Archive folder found**: Create an Archive folder in your Inbox or root level (see Setup Requirements section)
3. **Executable won't run**:
   - **Windows Security Blocking**: If Windows blocks the executable (common for downloaded files), the app will attempt to unblock it automatically
   - If automatic unblocking fails, right-click the .exe file → Properties → Check "Unblock" → Click OK
   - Ensure you're running as administrator if needed
   - Check Windows Defender isn't blocking the executable
   - Try running the batch file instead
4. **Config file issues**: The script will auto-create `config.json` if missing, but check for valid JSON syntax if errors occur
5. **Invalid JSON**: Check that your `config.json` file has valid JSON syntax - the script will show specific error details
6. **Scheduled task not running**:
   - Ensure Outlook is running when the task executes
   - Check task history in Task Scheduler
   - Verify the executable path is correct
   - Run the task manually first to test
7. **No log files created**: Check if the log directory path is accessible and writable

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

**Note**: This script is designed for personal use and should be tested thoroughly in your environment before production use. This is version 2.3.0 and is provided "as-is" with no planned updates unless critical issues are found.
