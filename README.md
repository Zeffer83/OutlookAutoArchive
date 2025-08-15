# Outlook Auto Archive

A professional PowerShell application that automatically archives emails older than a specified number of days from your Outlook Inbox to organized year/month folders. Features beautiful UI/UX with ASCII art banners, emojis, and professional styling.

**Author**: Ryan Zeffiretti  
**Version**: 2.9.5  
**License**: MIT

## ✨ Features

- **🎨 Beautiful User Interface**: Professional ASCII art banners, emojis, and color-coded sections
- **📧 Automatic Email Archiving**: Moves emails older than X days from Inbox to Archive folders
- **📁 Organized Structure**: Creates year/month folder hierarchy (e.g., `Archive\2025\2025-08`)
- **🛡️ Dry-Run Mode**: Test the application safely without actually moving emails
- **📝 Comprehensive Logging**: Detailed logs of all operations with timestamps
- **👥 Multi-Account Support**: Works with multiple Outlook accounts/stores
- **🔍 Smart Folder Detection**: Automatically finds Archive folders in various locations
- **🔄 Duplicate Prevention**: Handles duplicate emails intelligently
- **⚙️ Custom Skip Rules**: Built-in logic to skip specific emails (e.g., monitoring alerts)
- **✅ Outlook Status Check**: Automatically verifies Outlook is running before execution
- **🛠️ Enhanced Error Handling**: Improved logging and error recovery
- **⏰ Windows Task Scheduler**: Easy setup for automatic archiving
- **📱 Gmail Support**: Works with Gmail accounts using custom labels
- **💻 Console Compatibility**: ASCII-compatible display that works on all Windows systems
- **🎨 Professional UI**: Clean, readable interface with consistent styling across all terminals

## 🛡️ Safety Features

- **🛡️ Dry-Run Mode**: Test without making changes
- **🔄 Duplicate Prevention**: Avoids moving duplicate emails
- **🛠️ Error Handling**: Graceful handling of missing folders or permissions
- **📝 Backup Logging**: All operations are logged before execution
- **🔒 Windows Security**: Automatic unblocking of downloaded executables

## 📋 Prerequisites

- Windows 10/11 with PowerShell 5.1 or later
- Microsoft Outlook (desktop app) installed and configured
- Outlook COM Interop permissions

## 🔒 Windows Security

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

## 💻 Console Compatibility

### Universal Display Support

The application uses ASCII-compatible characters instead of Unicode emojis to ensure proper display across all Windows console environments:

- **✅ Consistent Display**: All text indicators use ASCII characters that display correctly everywhere
- **🖥️ Cross-Platform**: Works on all Windows systems regardless of console encoding
- **🎯 Professional Appearance**: Clean, readable text indicators like `[OK]`, `[ERROR]`, `[TIP]`
- **⚡ Better Performance**: No Unicode rendering issues or display glitches

### Text Indicators

The application uses these ASCII-compatible indicators:

- `[OK]` - Success messages and confirmations
- `[ERROR]` - Error messages and failures
- `[!]` - Warning messages and important notes
- `[TIP]` - Helpful tips and recommendations
- `[EMAIL]` - Email-related operations
- `[FOLDER]` - Folder operations
- `[SCHEDULE]` - Scheduling information
- `[STATS]` - Statistics and summaries

This ensures that all users see the same professional interface regardless of their Windows console configuration.

## 🚀 Installation

### Quick Start (Recommended)

1. **Download** `OutlookAutoArchive.exe` from the latest release
2. **Double-click** the executable to start the first-run setup
3. **Follow the guided setup** to configure your preferences
4. **Test in dry-run mode** to verify everything works
5. **Set up automatic scheduling** using the included setup executable

### Installation Details

The application automatically installs to `C:\Users\[YourUsername]\OutlookAutoArchive` on first run:

- ✅ **User-specific installation** (no admin permissions required)
- ✅ **Easy to find and manage**
- ✅ **Works with all Windows user types**
- ✅ **Automatic configuration file creation**

### Files Included

- **`OutlookAutoArchive.exe`** - Main application (recommended for all users)
- **`setup_task_scheduler.exe`** - Task scheduler setup executable (easy to use)
- **`setup_task_scheduler.ps1`** - Task scheduler setup script (for advanced users)
- **`config.json`** - Configuration file (auto-created on first run)

### 🎯 Why Use Executables?

**Benefits of using the .exe files:**

- ✅ **No PowerShell execution policy issues**
- ✅ **Double-click to run** - no command line needed
- ✅ **Works with Windows Task Scheduler** out of the box
- ✅ **Automatic Windows security unblocking**
- ✅ **Professional appearance** with proper metadata
- ✅ **Easy for non-technical users**

## ⚙️ Configuration

The application uses a `config.json` file for configuration. Edit this file to customize behavior:

```json
{
  "RetentionDays": 14,
  "DryRun": true,
  "LogPath": "./Logs",
  "GmailLabel": "OutlookArchive",
  "SkipRules": [
    {
      "Mailbox": "Your Mailbox Name",
      "Subjects": ["Subject Pattern 1", "Subject Pattern 2"]
    }
  ]
}
```

### Configuration Options

- **`RetentionDays`**: Number of days to keep emails in Inbox before archiving (default: 14)
- **`DryRun`**: When `true`, shows what would be moved without actually moving emails (default: true)
- **`LogPath`**: Directory where log files are stored (supports `%USERPROFILE%` variable)
- **`GmailLabel`**: Custom Gmail label name for archive folder (default: "OutlookArchive")
- **`SkipRules`**: Array of rules to skip specific emails by mailbox and subject patterns

## 🎯 How It Works

1. **🔗 Connects to Outlook**: Uses COM Interop to access Outlook
2. **📁 Finds Archive Folder**: Searches for Archive folder in multiple locations:
   - `Inbox\Archive`
   - Root-level `Archive`
   - Custom Gmail labels (configured via `GmailLabel`)
3. **📂 Creates Folder Structure**: Automatically creates year/month folders
4. **📧 Scans Inbox**: Processes all emails in the Inbox
5. **📦 Moves Old Emails**: Moves emails older than retention period to appropriate archive folder
6. **📝 Logs Everything**: Records all operations to timestamped log files

## 📁 Archive Structure

The application creates an organized folder structure:

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

## 🚀 Usage

### First Run Setup

1. **Double-click** `OutlookAutoArchive.exe`
2. **Follow the guided setup** to configure archive folders and scheduling
3. **Test in dry-run mode** to verify everything works
4. **Review logs** in the `Logs` folder
5. **Set up automatic scheduling** using the provided options

### Manual Execution

```powershell
# Run the executable (uses config.json automatically)
.\OutlookAutoArchive.exe
```

### Scheduled Execution

The first-run setup offers **daily scheduled archiving**:

> **💡 Design Philosophy**: We've simplified the scheduling to focus on reliability and system performance. Unlike continuous monitoring solutions that stay in memory, this approach runs once per day and then completely unloads, ensuring your system remains responsive and efficient.

#### 📅 Daily at Specific Time

- Runs once per day at your chosen time (e.g., 2:00 AM)
- Perfect for users who want predictable, scheduled archiving
- Gracefully skips runs when Outlook is not available
- Simple and reliable scheduling option
- **Memory efficient** - application unloads after each run, keeping your system responsive
- **Reliable performance** - no background processes consuming system resources

#### ⚙️ Manual Setup

Use `setup_task_scheduler.exe` to set up scheduling later:

```powershell
# Run as Administrator
.\setup_task_scheduler.exe
```

**Alternative**: Use Windows Task Scheduler directly:

1. Open **Task Scheduler** (search in Start menu)
2. Click **Create Basic Task**
3. Name: `Outlook Auto Archive`
4. Trigger: **Daily** at your preferred time
5. Action: **Start a program**
6. Program: `C:\Users\YourUsername\OutlookAutoArchive\OutlookAutoArchive.exe`
7. Check **Run with highest privileges**

## 📝 Logging

Logs are stored in: `Logs\` folder within your installation directory

Each run creates a timestamped log file: `ArchiveLog_YYYY-MM-DD_HH-mm-ss.txt`

Log entries include:

- Configuration settings
- Folder creation operations
- Email movement details
- Errors and warnings
- Completion timestamp

## 🎨 Customization

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

### Gmail Setup

For Gmail accounts, the application uses custom labels instead of folders:

1. **Enable IMAP** in Gmail settings
2. **Create a custom label** (e.g., "OutlookArchive")
3. **Check "Show in IMAP"** for the label
4. **Configure** `GmailLabel` in your `config.json`

## 🛠️ Troubleshooting

### Common Issues

1. **"Access Denied" errors**: Ensure Outlook is running and you have permissions
2. **No Archive folder found**: The application will help you create one during setup
3. **Executable won't run**:
   - Windows may block downloaded files - the app will attempt to unblock automatically
   - If automatic unblocking fails, right-click → Properties → Check "Unblock"
4. **Config file issues**: The app will auto-create `config.json` if missing
5. **Scheduled task not running**: Ensure Outlook is running when the task executes

### Debug Mode

Check the log files in the `Logs` folder for detailed error messages and operation history.

## 📄 License

This project is open source and available under the [MIT License](LICENSE).

**Author**: Ryan Zeffiretti

## ⚠️ Important Disclaimers

### 🔒 Data Safety Warning

**ALWAYS BACKUP YOUR DATA BEFORE USE!** While this application includes safety features like dry-run mode and comprehensive logging, it's your responsibility to ensure you have proper backups of your email data before using this tool.

### 🛡️ No Warranty

This software is provided "AS IS" without warranty of any kind. The author makes no representations or warranties about the accuracy, reliability, completeness, or suitability of this software for any purpose.

### 📋 User Responsibility

By using this software, you acknowledge that:

- You have backed up your email data before use
- You understand the risks involved in email operations
- You accept full responsibility for any consequences
- You will test the software in dry-run mode first

## 🆘 Support

If you encounter issues:

1. **Check the log files** for detailed error messages
2. **Ensure Outlook is properly configured**
3. **Verify you have necessary permissions**
4. **Test with dry-run mode enabled**
5. **Review the troubleshooting section above**

---

**Note**: This application is designed for personal use and should be tested thoroughly in your environment before production use. Version 2.9.5 includes enhanced UI/UX with beautiful styling, professional interface design, and improved console compatibility.
