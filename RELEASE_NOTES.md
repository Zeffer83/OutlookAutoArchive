# Release Notes - Version 1.2.0

## 🎉 What's New in Version 1.2.0

### ✨ Major Improvements
- **Executable Version**: Now includes `OutlookAutoArchive.exe` for easy deployment
- **User-Friendly Batch File**: Added `Run_OutlookAutoArchive.bat` for simple double-click execution
- **No PowerShell Required**: Users can run the tool without dealing with PowerShell execution policies
- **Enhanced Task Scheduler Support**: Easier integration with Windows Task Scheduler

### 🚀 Key Features
- **Double-Click Execution**: Simply double-click the batch file to run
- **Automatic Configuration**: Still uses `config.json` for all settings
- **Safe Defaults**: Executable starts in dry-run mode by default
- **Cross-Platform Compatibility**: Works on any Windows system with Outlook

### 📋 What's Included
- `OutlookAutoArchive.exe` - The main executable
- `Run_OutlookAutoArchive.bat` - User-friendly batch file
- `OutlookAutoArchive.ps1` - Original PowerShell script
- `config.example.json` - Example configuration
- `README.md` - Complete documentation

### 🔧 Installation
1. Download and extract the files
2. Create an Archive folder in Outlook (see README for details)
3. Double-click `Run_OutlookAutoArchive.bat` to run
4. Edit `config.json` to customize settings

### ⚠️ Important Notes
- **First Run**: The script will create a `config.json` file with safe defaults
- **Dry-Run Mode**: Default configuration starts in dry-run mode for safety
- **Outlook Required**: Microsoft Outlook must be installed and running
- **Backup Recommended**: Always backup your emails before first use

### 🐛 Bug Fixes & Improvements
- Enhanced error handling for missing configuration files
- Improved user experience with clear console messages
- Better documentation with step-by-step instructions
- Simplified installation process

---

## Previous Versions

### Version 1.1.0
- External configuration via `config.json`
- Enhanced skip rules system
- Improved Gmail label support
- Auto-configuration creation

### Version 1.0.0
- Initial release with core archiving functionality
- Dry-run mode and comprehensive logging
- Multi-account support and smart folder detection

---

**Author**: Ryan Zeffiretti  
**License**: MIT  
**Support**: This is an "as-is" script with limited maintenance
