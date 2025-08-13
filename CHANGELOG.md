# Changelog

All notable changes to the Outlook Auto Archive Script will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.9.0] - 2025-08-13

**Author**: Ryan Zeffiretti

### Added

- **Integrated Outlook Status Check**: Enhanced Outlook status checking with better user feedback and guidance
- **Convenience Batch File Generation**: Option to create convenience .bat files during first-run setup
- **Single-File Download Experience**: Simplified download with only essential files included
- **Interactive Batch File Creation**: Users can choose which convenience scripts to create

### Changed

- **Simplified File Structure**: Removed unnecessary .bat files from distribution
- **Enhanced User Experience**: Better error messages and user guidance for Outlook status issues
- **Streamlined Setup Process**: All functionality now integrated into the main executable
- **Cleaner Repository**: Reduced file clutter and improved organization

### Removed

- **Standalone Batch Files**: Removed `Run_OutlookAutoArchive.bat`, `Run_OutlookAutoArchive_WithCheck.bat`, and `Setup_Archive_Folders.bat` from distribution
- **File Clutter**: Eliminated unnecessary wrapper files

### Technical Details

- **Integrated Functionality**: All batch file functionality now integrated into main executable
- **Dynamic File Generation**: Convenience .bat files generated on-demand during first run
- **Improved Error Handling**: Better user feedback when Outlook is not running
- **Simplified Distribution**: Single executable download with optional convenience files

## [1.8.0] - 2025-08-13

**Author**: Ryan Zeffiretti

### Added

- **Interactive Installation Location Setup**: New guided installation setup during first-run configuration
- **Multiple Installation Options**: Users can choose between Program Files, User Documents, custom location, or current location
- **Automatic File Migration**: Automatically copies all application files to the chosen installation location
- **Path Validation**: Validates custom installation paths to ensure they are valid
- **Installation Feedback**: Provides detailed feedback during the installation process
- **File Management**: Comprehensive file copying with error handling and rollback

### Changed

- **Enhanced First-Run Experience**: First-run setup now includes complete installation location configuration
- **User Experience**: Streamlined installation process with guided location selection
- **File Organization**: Better organization of application files in chosen installation directory
- **Documentation**: Updated to reflect new installation capabilities

### Technical Details

- **Installation Options**: Program Files (system-wide), User Documents (user-specific), custom location, or current location
- **File Migration**: Copies all necessary files including executables, scripts, documentation, and configuration files
- **Path Handling**: Intelligent path detection and validation for custom installation locations
- **Error Recovery**: Graceful error handling with fallback to current location if installation fails
- **Directory Creation**: Automatic creation of installation directories if they don't exist

## [1.7.0] - 2025-08-13

**Author**: Ryan Zeffiretti

### Added

- **Interactive Scheduled Task Setup**: New guided scheduling setup during first-run configuration
- **Multiple Scheduling Options**: Users can choose between daily time-based scheduling, Outlook startup integration, or manual setup
- **Automatic Task Creation**: Automatically creates Windows Task Scheduler tasks based on user preferences
- **Time Input Validation**: Validates user time input in 24-hour format for daily scheduling
- **Fallback Mechanisms**: Graceful fallback to PowerShell script if executable is not found
- **Error Handling**: Comprehensive error handling for task creation with helpful fallback instructions

### Changed

- **Enhanced First-Run Experience**: First-run setup now includes complete scheduling configuration
- **User Experience**: Streamlined setup process with guided scheduling options
- **Task Scheduler Integration**: Seamless integration with Windows Task Scheduler
- **Documentation**: Updated to reflect new scheduling capabilities

### Technical Details

- **Daily Scheduling**: Creates daily tasks at user-specified times using schtasks command
- **Startup Integration**: Option to use existing Setup_OutlookStartup_Task.ps1 or create basic startup task
- **Path Detection**: Intelligent detection of executable vs PowerShell script paths
- **Task Naming**: Consistent task naming conventions for easy identification
- **Validation**: Time format validation and user input sanitization

## [1.6.0] - 2025-08-13

**Author**: Ryan Zeffiretti

### Added

- **Interactive First-Run Setup**: New guided setup process that automatically detects email accounts and configures archive folders/labels.
- **Smart Account Detection**: Automatically detects Gmail accounts (including @gmail.co.uk) and regular email accounts.
- **Configurable Retention Period**: Interactive setup allows users to set their preferred retention period (default: 14 days).
- **Gmail Label Configuration**: Interactive setup for Gmail label names with validation (prevents "Archive" which is not allowed).
- **Archive Folder Path Storage**: Stores detected archive folder paths in config.json for faster future access.
- **Backward Compatibility**: Falls back to folder search if stored paths are not found.

### Changed

- **Performance Improvement**: Archive folder detection is now much faster using stored paths instead of searching every time.
- **User Experience**: First-time users get a guided setup experience instead of manual configuration.
- **Configuration Management**: Archive folder paths are automatically stored and managed in config.json.
- **Gmail Support**: Enhanced Gmail detection including @gmail.co.uk domains.

### Technical Details

- **Path Storage Format**: Archive folder paths stored as "Type:Name" (e.g., "GmailLabel:OutlookArchive", "Root:Archive", "Inbox:Archive").
- **Fallback Mechanism**: If stored paths fail, script falls back to original search method for compatibility.
- **Config Structure**: Added `OnFirstRun` and `ArchiveFolders` fields to config.json.
- **Validation**: Gmail label names validated to prevent invalid characters.

## [1.5.0] - 2025-08-13

**Author**: Ryan Zeffiretti

### Added

- **Complete Folder Detection System**: Fixed critical issue where script couldn't find archive folders in Outlook accounts.
- **Multi-Account Support**: Script now properly detects and processes all email accounts (Gmail, Outlook, Exchange, etc.).
- **Gmail Label Support**: Full support for Gmail labels as archive folders (e.g., "OutlookArchive").
- **Comprehensive Testing**: Thorough testing completed with 1,000+ emails processed across multiple accounts.

### Fixed

- **Outlook Folder Access**: Completely rewrote folder detection logic to use `$namespace.Folders` instead of `$namespace.Stores`.
- **Gmail Label Detection**: Fixed Gmail label access by using direct folder enumeration through namespace.
- **Account Processing**: Script now correctly processes all account types and finds their respective archive folders.
- **Skip Rules Functionality**: Verified skip rules work correctly across all account types.

### Changed

- **Architecture**: Changed from store-based to account-based processing for better compatibility.
- **Folder Detection**: Improved `Get-ArchiveFolder` function to work with account objects instead of store objects.
- **Error Handling**: Enhanced error handling for folder access and account processing.
- **Documentation**: Updated README to reflect successful testing and working status.

### Technical Details

- **Namespace Access**: Switched from `$namespace.Stores` to `$namespace.Folders` for proper account enumeration.
- **Account Processing**: Updated all functions to work with account objects (`$account`) instead of store objects (`$store`).
- **Folder Detection**: Improved logic to find archive folders in Inbox, root level, and Gmail labels.
- **Testing Results**: Successfully tested with Gmail accounts (using labels), regular email accounts (using folders), and skip rules.

## [1.4.0] - 2025-08-13

**Author**: Ryan Zeffiretti

### Fixed

- **Path Binding Errors**: Resolved "cannot bind argument to parameter path" errors
- **Null Path Issues**: Fixed null path binding errors in logging setup
- **JSON Parsing**: Improved config.json handling and error recovery
- **Executable Console Mode**: Fixed popup windows by using proper console mode
- **Log Path Processing**: Enhanced handling of environment variables and backslashes

### Changed

- **Enhanced Error Handling**: More robust error handling throughout the script
- **Improved Logging**: Better logging system with safe file writing
- **Path Resolution**: Better path handling for both script and executable modes
- **User Experience**: Emphasized executable usage for end users
- **Documentation**: Updated to clarify executable vs source code usage

### Technical Details

- **Safe Logging**: New `Write-Log` function with proper error handling
- **Path Detection**: Improved script directory detection for executables
- **JSON Validation**: Better error messages for configuration issues
- **Console Mode**: Proper console application build for better user experience

## [1.3.0] - 2025-08-13

**Author**: Ryan Zeffiretti

### Added

- **Outlook Status Check**: Automatic verification that Outlook is running before execution
- **Enhanced Batch File**: `Run_OutlookAutoArchive_WithCheck.bat` with Outlook status verification
- **Startup Task Setup**: `Setup_OutlookStartup_Task.ps1` for creating Outlook startup tasks
- **Improved Error Handling**: Better logging and error recovery for path issues
- **Outlook Startup Integration**: Task that waits for Outlook to start before running archive

### Changed

- **Enhanced Safety**: Script now checks Outlook status before attempting operations
- **Better Logging**: Improved error handling and logging with null checks
- **Scheduled Task Options**: Added method to run when Outlook starts automatically
- **Documentation**: Updated with new features and improved setup instructions

### Technical Details

- **Outlook Detection**: Uses `Get-Process` to verify Outlook is running
- **Path Handling**: Improved error handling for log path processing
- **Task Scheduler**: New PowerShell script for creating Outlook startup tasks
- **Error Recovery**: Graceful handling of missing paths and configuration issues

## [1.2.0] - 2025-08-13

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

## [1.1.0] - 2025-08-13

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

## [1.0.0] - 2025-08-13

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
- **Organized Structure**: Creates year/month hierarchy (e.g., `Archive\2025\2025-08`)
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

- **1.4.0**: Bug fixes and improved error handling (2025-08-13)
- **1.3.0**: Outlook status check and startup integration (2025-08-13)
- **1.2.0**: Executable version and enhanced user experience (2025-08-13)
- **1.1.0**: External configuration and enhanced skip rules (2025-08-13)
- **1.0.0**: Initial release with core functionality (2025-08-13)
