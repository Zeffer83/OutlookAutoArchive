# Changelog

All notable changes to the Outlook Auto Archive Script will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2024-12-19

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

---

## Version History

- **1.0.0**: Initial release with core functionality (2024-12-19)
