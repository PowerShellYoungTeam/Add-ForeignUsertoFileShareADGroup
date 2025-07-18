#
# Module: ADGroupTool
# Author: Steven Wight with GitHub Copilot
# Description: Enhanced AD Group Management Tool for adding users from foreign domains to Active Directory groups
# Compatible with: Windows Server 2012 R2+ and PowerShell 5.1+
#

# Get public and private function definition files
$Public = @(Get-ChildItem -Path $PSScriptRoot\Functions\*.ps1 -ErrorAction SilentlyContinue)
$GUI = @(Get-ChildItem -Path $PSScriptRoot\GUI\*.ps1 -ErrorAction SilentlyContinue)
$Install = @(Get-ChildItem -Path $PSScriptRoot\Install\*.ps1 -ErrorAction SilentlyContinue)

# Dot source the files
foreach ($import in @($Public + $GUI + $Install)) {
    try {
        Write-Verbose "Importing $($import.FullName)"
        . $import.FullName
    }
    catch {
        Write-Error "Failed to import function $($import.FullName): $($_.Exception.Message)"
    }
}

# Export public functions
Export-ModuleMember -Function $Public.BaseName
Export-ModuleMember -Function 'Start-ADGroupToolGUI'
Export-ModuleMember -Function 'Install-ADGroupTool'

# Create aliases for backward compatibility
New-Alias -Name 'Add-ForeignUserToADGroup' -Value 'Invoke-ADGroupOperation' -Force
New-Alias -Name 'Start-ADGroupGUI' -Value 'Start-ADGroupToolGUI' -Force

Export-ModuleMember -Alias 'Add-ForeignUserToADGroup'
Export-ModuleMember -Alias 'Start-ADGroupGUI'

# Module variables
$Script:ModuleRoot = $PSScriptRoot
$Script:ModuleVersion = (Import-PowerShellDataFile -Path "$PSScriptRoot\ADGroupTool.psd1").ModuleVersion

# Initialize module
Write-Verbose "ADGroupTool module v$Script:ModuleVersion loaded successfully"
Write-Verbose "Module root: $Script:ModuleRoot"

# Check prerequisites on module import
try {
    # Check PowerShell version
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        Write-Warning "PowerShell 5.1 or later is required. Current version: $($PSVersionTable.PSVersion)"
    }

    # Check for Active Directory module
    if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
        Write-Warning "Active Directory PowerShell module is not available. Please install RSAT tools."
        Write-Host "To install: Enable-WindowsOptionalFeature -Online -FeatureName RSATTools-AD-PowerShell" -ForegroundColor Yellow
    }

    # Check platform compatibility
    if ($PSVersionTable.Platform -and $PSVersionTable.Platform -ne "Win32NT") {
        Write-Warning "This module is designed for Windows platforms only."
    }
}
catch {
    Write-Warning "Error during module initialization: $($_.Exception.Message)"
}

# Module cleanup
$MyInvocation.MyCommand.ScriptBlock.Module.OnRemove = {
    Write-Verbose "ADGroupTool module is being removed"
}
