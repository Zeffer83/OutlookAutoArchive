# Changelog

All notable changes to the Outlook Auto Archive Script will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.2.0] - 2024-12-19

**Author**: Ryan Zeffiretti

### Added
- **Executable Version**: Created `OutlookAutoArchive.exe` for easy deployment
- **Batch File**: Added `Run_OutlookAutoArchive.bat` for user-friendly execution
- **Enhanced User Experience**: No PowerShell execution policy issues
- **Task Scheduler Support**: Easy integration with Windows Task Scheduler
- **Double-Click Execution**: Users can run by double-clicking the batch file

### Changed
- **Installation**: Simplified with executable and batch file options
- **Documentation**: Updated README with executable usage instructions
- **Version Management**: Updated to version 1.2.0
- **GitHub Preparation**: Project ready for GitHub release

### Technical Details
- **Executable**: Created using PS2EXE tool
- **Compatibility**: Works on Windows with Outlook installed
- **Configuration**: Still uses `config.json` for all settings

## [1.1.0] - 2024-12-19

**Author**: Ryan Zeffiretti

### Changed
- **External Configuration**: Moved from hardcoded settings to `config.json` file
- **Enhanced Skip Rules**: Configurable skip rules via JSON configuration
- **Improved Gmail Support**: Configurable Gmail label support
- **Better Error Handling**: Enhanced error handling for missing config files

### Added
- **JSON Configuration**: All settings now configurable via `config.json`
- **Flexible Skip Rules**: Support for multiple mailbox-specific skip rules
- **Environment Variable Support**: `%USERPROFILE%` variable support in log paths
- **Config Validation**: Basic JSON validation and error reporting

## [1.0.0] - 2024-12-19

**Author**: Ryan Zeffiretti

### Added
- Initial release of Outlook Auto Archive Script
- Automatic email archiving with configurable retention period
- Dry-run mode for safe testing
- Comprehensive logging system
- Multi-account support
- Smart archive folder detection
- Duplicate email prevention
- Custom skip rules for specific emails
- Year/month folder organization structure

### Features
- **Archive Folder Detection**: Automatically finds Archive folders in multiple locations:
  - `Inbox\Archive`
  - Root-level `Archive`
  - Gmail-style `OutlookArchive`
- **Organized Structure**: Creates year/month hierarchy (e.g., `Archive\2024\2024-12`)
- **Safety Features**: Dry-run mode, duplicate prevention, comprehensive logging
- **Custom Skip Rules**: Built-in logic to skip monitoring alerts and other specific emails

### Technical Details
- PowerShell script using Outlook COM Interop
- Configurable retention period (default: 14 days)
- Timestamped log files in user documents folder
- Error handling for missing folders and permissions
- Support for multiple Outlook accounts/stores

### Notes
- This is an "as-is" script created for personal use
- No planned updates unless critical issues are found

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

---

## Version History

- **1.2.0**: Executable version and enhanced user experience (2024-12-19)
- **1.1.0**: External configuration and enhanced skip rules (2024-12-19)
- **1.0.0**: Initial release with core functionality (2024-12-19)
