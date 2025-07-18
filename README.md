# Enhanced AD Group Management Tool v2.0

A comprehensive PowerShell solution for adding users from foreign domains to Active Directory groups with advanced error handling, validation, and a modern GUI interface.

## 🚀 New Features in v2.0

### Core Script Enhancements
- **Comprehensive Validation**: Pre-flight checks for CSV schema, domain connectivity, and user/group existence
- **Retry Logic**: Configurable retry attempts with exponential backoff for transient failures
- **Performance Monitoring**: Detailed timing and performance metrics
- **Enhanced Logging**: Structured logging with operation IDs and comprehensive audit trails
- **Parallel Processing**: Optional parallel execution for large datasets
- **Security Improvements**: Better credential handling and validation

### GUI Controller Enhancements
- **Modern Tabbed Interface**: Organized into Configuration, Advanced, Email, and Logs tabs
- **Real-time Validation**: Live input validation with visual feedback
- **Configuration Persistence**: Save and load user preferences
- **Progress Monitoring**: Real-time progress bars and status updates
- **Enhanced Error Handling**: Detailed error messages with context
- **CSV Preview**: Preview CSV data before execution
- **Log Viewer**: Real-time log viewing and export capabilities

## 📋 Prerequisites

- Windows Server 2012 R2 or later / Windows 10/11
- Windows PowerShell 5.1 or later
- Active Directory PowerShell module (RSAT Tools)
- Appropriate domain permissions for user and group management
- Network connectivity to source and target domains

## 🏗️ Installation

1. **Download the files** to a local directory:
   - `Add-ForeignUsertoFileShareADGroup.ps1` (Core script)
   - `Add-ForeignUsertoFileShareADGroupGUIController.ps1` (GUI interface)
   - `Sample_UserGroups.csv` (Example CSV format)

2. **Install Prerequisites**:
   ```powershell
   # Install RSAT Tools (if not already installed)
   Enable-WindowsOptionalFeature -Online -FeatureName RSATTools-AD-PowerShell

   # Or on Windows Server:
   Install-WindowsFeature RSAT-AD-PowerShell
   ```

3. **Set Execution Policy** (if needed):
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

## 📊 CSV File Format

Your input CSV file must contain the following columns:

| Column | Description | Example |
|--------|-------------|---------|
| SourceDomain | Domain where the user exists | `contoso.com` |
| SourceUser | Username (without domain) | `jdoe` |
| TargetDomain | Domain where the group exists | `fabrikam.com` |
| TargetGroup | Group name to add user to | `FileShare_Readers` |

### Sample CSV Content:
```csv
SourceDomain,SourceUser,TargetDomain,TargetGroup
contoso.com,jdoe,fabrikam.com,FileShare_Readers
contoso.com,asmith,fabrikam.com,FileShare_Writers
external.com,consultant1,fabrikam.com,Temp_Access
```

## 🖥️ Usage

### GUI Method (Recommended)
1. **Launch the GUI**:
   ```powershell
   .\Add-ForeignUsertoFileShareADGroupGUIController.ps1
   ```

2. **Configure Settings**:
   - **Configuration Tab**: Set script path, CSV file, and output folder
   - **Advanced Tab**: Configure retry logic and parallel processing
   - **Email Tab**: Set up email notifications
   - **Logs Tab**: Monitor real-time execution

3. **Execute**:
   - Click "Run Script" to start execution
   - Monitor progress in the Logs tab
   - Review results in the output folder

### Command Line Method
```powershell
.\Add-ForeignUsertoFileShareADGroup.ps1 -InputCsvPath "C:\Path\To\Users.csv" -OutputFolderPath "C:\Output" -Test
```

### Parameters

#### Core Script Parameters
| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `InputCsvPath` | String | Yes | Path to CSV file with user/group data |
| `OutputFolderPath` | String | Yes | Output directory for logs |
| `Test` | Switch | No | Run in test mode (WhatIf) |
| `ForeignAdminCreds` | PSCredential | No | Credentials for foreign domain |
| `MaxRetries` | Int | No | Maximum retry attempts (1-10, default: 3) |
| `RetryDelaySeconds` | Int | No | Delay between retries (1-60, default: 5) |
| `ParallelProcessing` | Switch | No | Enable parallel processing |

## 🔧 Advanced Features

### Retry Logic
The script automatically retries failed operations with configurable settings:
- Default: 3 retries with 5-second delays
- Exponential backoff for certain error types
- Detailed logging of retry attempts

### Performance Monitoring
- Operation-level timing
- Performance metrics reporting
- Batch processing statistics
- Memory usage monitoring

### Security Enhancements
- Credential validation before processing
- Secure credential storage options
- Audit logging for compliance
- Permission verification

### Error Handling
- Comprehensive error categorization
- Detailed error context
- Recovery suggestions
- Graceful failure handling

## 📈 Output Files

The script generates several output files:

1. **Log File**: `ADUserToGroupLog_[MODE]_[TIMESTAMP]_[COMPUTER]_[USER].csv`
   - Detailed operation results
   - Performance metrics
   - Error details

2. **Transcript**: `Transcript_[MODE]_[TIMESTAMP]_[COMPUTER]_[USER].txt`
   - Complete execution transcript
   - Verbose output

3. **Summary**: `Summary_[MODE]_[TIMESTAMP]_[COMPUTER]_[USER].txt`
   - Executive summary
   - Performance statistics
   - File locations

### Log File Columns
| Column | Description |
|--------|-------------|
| OperationId | Unique identifier for the operation |
| Timestamp | When the operation occurred |
| SourceUser | Source user account |
| SourceDomain | Source domain |
| SourceUserDN | Full distinguished name |
| TargetGroup | Target group name |
| TargetDomain | Target domain |
| Status | Success/Error/Skipped |
| Message | Detailed result message |
| DurationSeconds | Operation duration |
| ProcessedBy | User who ran the script |
| ComputerName | Computer where script ran |
| TestMode | Whether in test mode |

## 🔧 Configuration

### GUI Configuration Persistence
The GUI automatically saves and loads configuration including:
- File paths and folders
- Email settings
- Advanced options
- User preferences

Configuration is stored in: `%APPDATA%\ADGroupScriptGUI_Config.xml`

### Email Notifications
Configure automatic email notifications:
- SMTP server settings
- Recipient configuration
- Custom message templates
- Attachment of log files

## 🛠️ Troubleshooting

### Common Issues

1. **"ActiveDirectory module not found"**
   - Install RSAT Tools or run on Domain Controller
   ```powershell
   Install-WindowsFeature RSAT-AD-PowerShell
   ```

2. **"Access Denied" errors**
   - Verify domain credentials
   - Check user permissions in target domain
   - Ensure network connectivity

3. **CSV validation errors**
   - Check column names match exactly
   - Remove empty rows
   - Verify UTF-8 encoding

4. **Performance issues**
   - Enable parallel processing for large datasets
   - Increase retry delays for slow networks
   - Check domain controller performance

### Validation Checklist
✅ PowerShell 5.1+ installed
✅ Active Directory module available
✅ Network connectivity to all domains
✅ Appropriate permissions granted
✅ CSV file format correct
✅ Output folder writable

## 📝 Examples

### Basic Usage
```powershell
# Test mode execution
.\Add-ForeignUsertoFileShareADGroup.ps1 -InputCsvPath ".\users.csv" -OutputFolderPath ".\output" -Test

# Live execution with custom retry settings
.\Add-ForeignUsertoFileShareADGroup.ps1 -InputCsvPath ".\users.csv" -OutputFolderPath ".\output" -MaxRetries 5 -RetryDelaySeconds 10

# Parallel processing for large datasets
.\Add-ForeignUsertoFileShareADGroup.ps1 -InputCsvPath ".\large_users.csv" -OutputFolderPath ".\output" -ParallelProcessing
```

### GUI Examples
1. **Quick Setup**: Use default paths and run in test mode
2. **Production Run**: Configure all settings, test email, then execute
3. **Batch Processing**: Load large CSV, enable parallel processing
4. **Monitoring**: Use logs tab for real-time monitoring

## 🔐 Security Considerations

- Store credentials securely
- Use least-privilege accounts
- Monitor audit logs
- Regular permission reviews
- Network security (encrypted connections)

## 📋 Version History

### v2.0 (Current)
- Complete rewrite with enhanced error handling
- Modern GUI with tabbed interface
- Real-time validation and monitoring
- Configuration persistence
- Parallel processing support
- Comprehensive logging

### v1.0
- Basic script functionality
- Simple GUI interface
- Basic error handling

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 👥 Authors

- Steven Wight with GitHub Copilot
- Enhanced for enterprise use

## 🆘 Support

For support and questions:
1. Check the troubleshooting section
2. Review log files for detailed error information
3. Ensure all prerequisites are met
4. Contact your system administrator

---

**Note**: Always test in a non-production environment first. Use test mode to validate operations before making live changes.
