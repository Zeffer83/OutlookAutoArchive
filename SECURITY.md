# Security Policy

**Author**: Ryan Zeffiretti

## Reporting Security Issues

This is an "as-is" script with limited maintenance. If you discover a security vulnerability:

1. **DO NOT** create a public GitHub issue
2. Email your findings to: **rzeffiretti@gmail.com**
3. Include a description of the issue and steps to reproduce

## Security Considerations

This script interacts with Outlook and email data. Please be aware of:

- **Email Data Access**: The script reads and moves email messages
- **Outlook Permissions**: Requires COM Interop access to Outlook
- **Log Files**: Contains information about email operations
- **Local Execution**: Runs locally on your machine

## Security Features

- **Dry-Run Mode**: Test without making changes
- **Comprehensive Logging**: Track all operations
- **Error Handling**: Graceful failure without data loss

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

**Note**: This is an "as-is" script created for personal use. Response times for security issues may vary.
