# Security Policy

**Author**: Ryan Zeffiretti

## Supported Versions

Use this section to tell people about which versions of your project are currently being supported with security updates.

| Version | Supported          |
| ------- | ------------------ |
| 1.0.x   | :white_check_mark: |

## Reporting a Vulnerability

We take security vulnerabilities seriously. If you discover a security vulnerability in the Outlook Auto Archive Script, please follow these steps:

### How to Report

1. **DO NOT** create a public GitHub issue for security vulnerabilities
2. Email your findings to: rzeffiretti@gmail.com
3. Include detailed information about the vulnerability
4. Provide steps to reproduce the issue
5. Suggest potential fixes if possible

### What to Include

When reporting a vulnerability, please include:

- Description of the vulnerability
- Steps to reproduce
- Potential impact
- Suggested fix (if any)
- Your contact information

### Response Timeline

- **Initial Response**: Within 48 hours
- **Status Update**: Within 1 week
- **Resolution**: As quickly as possible, typically within 30 days

### Security Considerations

This script interacts with Outlook and email data. Please be aware of:

1. **Email Data Access**: The script reads and moves email messages
2. **Outlook Permissions**: Requires COM Interop access to Outlook
3. **Log Files**: Contains information about email operations
4. **Local Execution**: Runs locally on your machine

### ⚠️ Important Disclaimers

#### 🔒 Data Safety Warning

**ALWAYS BACKUP YOUR DATA BEFORE USE!** While this script includes safety features like dry-run mode and comprehensive logging, it's your responsibility to ensure you have proper backups of your email data before using this tool.

#### 🛡️ No Warranty

This software is provided "AS IS" without warranty of any kind. The author makes no representations or warranties about the accuracy, reliability, completeness, or suitability of this software for any purpose.

#### 🛡️ Limitation of Liability

The author shall not be liable for any direct, indirect, incidental, special, consequential, or punitive damages, including but not limited to:

- Loss of data or emails
- Email corruption or deletion
- System corruption
- Any other damages arising from the use of this software

#### 📋 User Responsibility

By using this software, you acknowledge that:

- You have backed up your email data before use
- You understand the risks involved in email operations
- You accept full responsibility for any consequences
- You will test the software in dry-run mode first
- You have read and understood all disclaimers

### Best Practices

- Always test in dry-run mode first
- Review log files for sensitive information
- Keep the script updated
- Use appropriate file permissions
- Don't share log files containing sensitive data

### Disclosure Policy

- Vulnerabilities will be disclosed publicly after fixes are available
- Credit will be given to reporters who follow responsible disclosure
- Security updates will be released as patch versions

## Security Features

The script includes several security features:

- **Dry-Run Mode**: Test without making changes
- **Logging**: Track all operations for audit purposes
- **Error Handling**: Graceful failure without data loss
- **Permission Checks**: Validates access before operations

## Contact

For security-related issues, please contact: rzeffiretti@gmail.com

---

**Note**: This is an "as-is" script with limited maintenance. Response times may vary.
