# ADGroupTool Installation Guide

## Quick Start

1. **Import the installation function**:
   ```powershell
   . .\ADGroupTool\Install\Install-ADGroupTool.ps1
   ```

2. **Install the module**:
   ```powershell
   # For current user with desktop shortcut
   Install-ADGroupTool -CreateDesktopShortcut

   # For all users (requires admin)
   Install-ADGroupTool -Scope AllUsers -InstallRSAT -CreateDesktopShortcut
   ```

3. **Start using the tool**:
   ```powershell
   Import-Module ADGroupTool
   Start-ADGroupToolGUI
   ```

## Module Structure

The tool is now organized as a proper PowerShell module:

```
ADGroupTool/
├── ADGroupTool.psd1                    # Module manifest
├── ADGroupTool.psm1                    # Main module file
├── Functions/                          # Core functions
│   ├── Add-ADUserToGroup.ps1          # Main processing function
│   ├── Test-CSVSchema.ps1             # CSV validation
│   ├── Test-Credential.ps1            # Credential validation
│   ├── Test-DomainConnectivity.ps1    # Domain connectivity testing
│   ├── Invoke-WithRetry.ps1           # Retry logic
│   └── Invoke-ADGroupOperation.ps1    # Main wrapper function
├── GUI/
│   └── Start-ADGroupToolGUI.ps1       # GUI interface
├── Install/
│   └── Install-ADGroupTool.ps1        # Installation function
└── Data/
    └── Sample_UserGroups.csv           # Sample CSV file
```

## Key Improvements

### ✅ Modular Architecture
- Professional PowerShell module with proper manifest
- Functions separated into logical files
- Easy installation and management

### ✅ Enhanced Validation
- Pre-flight credential validation
- Domain connectivity testing
- Comprehensive CSV schema validation

### ✅ Better Error Handling
- Categorized error types
- Retry logic with exponential backoff
- Detailed logging with operation IDs

### ✅ Modern GUI
- Tabbed interface (Configuration, Advanced, Logs)
- Real-time validation and feedback
- Configuration persistence

### ✅ Improved Security
- Better credential handling
- Validation before operations
- Comprehensive audit trails

## Testing

Run the test script to validate the module:

```powershell
.\Test-ADGroupToolModule.ps1
```

This will verify:
- Module imports correctly
- All functions are available
- CSV validation works
- Help documentation is accessible

## Backwards Compatibility

The module maintains backwards compatibility through:
- Aliases: `Add-ForeignUserToADGroup` → `Invoke-ADGroupOperation`
- Aliases: `Start-ADGroupGUI` → `Start-ADGroupToolGUI`
- Same parameter names and functionality

## Next Steps

1. Test the module in your environment
2. Review the enhanced validation features
3. Customize configuration as needed
4. Deploy to production servers
