# Enhanced AD Group Management Tool - Installation Script
# Author: Steven Wight with GitHub Copilot
# Compatible with Windows PowerShell 5.1

param(
    [Parameter(Mandatory = $false)]
    [string]$InstallPath = "C:\ADGroupTool",

    [Parameter(Mandatory = $false)]
    [switch]$CreateDesktopShortcut,

    [Parameter(Mandatory = $false)]
    [switch]$InstallRSAT
)

Write-Host "=== Enhanced AD Group Management Tool - Installation ===" -ForegroundColor Green
Write-Host ""

# Check if running as Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin -and $InstallRSAT) {
    Write-Warning "Administrator privileges required for RSAT installation. Please run as Administrator or skip RSAT installation."
    $InstallRSAT = $false
}

try {
    # Create installation directory
    if (-not (Test-Path $InstallPath)) {
        Write-Host "Creating installation directory: $InstallPath" -ForegroundColor Yellow
        New-Item -Path $InstallPath -ItemType Directory -Force | Out-Null
    }

    # Check prerequisites
    Write-Host "Checking prerequisites..." -ForegroundColor Yellow

    # PowerShell version
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        throw "PowerShell 5.1 or later is required. Current version: $($PSVersionTable.PSVersion)"
    }
    Write-Host "✓ PowerShell version: $($PSVersionTable.PSVersion)" -ForegroundColor Green

    # Windows version
    if ($PSVersionTable.Platform -and $PSVersionTable.Platform -ne "Win32NT") {
        throw "This tool requires Windows PowerShell on Windows."
    }
    Write-Host "✓ Windows platform detected" -ForegroundColor Green

    # Active Directory module
    $adModule = Get-Module -Name ActiveDirectory -ListAvailable
    if (-not $adModule) {
        Write-Warning "Active Directory PowerShell module not found."
        if ($InstallRSAT) {
            Write-Host "Installing RSAT Tools..." -ForegroundColor Yellow

            # Check Windows version for appropriate installation method
            $osVersion = [System.Environment]::OSVersion.Version
            if ($osVersion.Major -eq 10) {
                # Windows 10/Server 2016+
                try {
                    Enable-WindowsOptionalFeature -Online -FeatureName RSATTools-AD-PowerShell -All -NoRestart
                    Write-Host "✓ RSAT Tools installation initiated" -ForegroundColor Green
                }
                catch {
                    Write-Warning "Failed to install RSAT via Windows Features. Trying alternative method..."
                    # Try DISM method
                    & dism /online /enable-feature /featurename:RSATTools-AD-PowerShell /all
                }
            }
            else {
                # Windows Server
                try {
                    Install-WindowsFeature RSAT-AD-PowerShell
                    Write-Host "✓ RSAT Tools installed" -ForegroundColor Green
                }
                catch {
                    Write-Warning "Failed to install RSAT: $($_.Exception.Message)"
                }
            }
        }
        else {
            Write-Host "⚠ Please install RSAT Tools manually:" -ForegroundColor Yellow
            Write-Host "  Windows 10/11: Settings > Apps > Optional Features > RSAT" -ForegroundColor Yellow
            Write-Host "  Windows Server: Install-WindowsFeature RSAT-AD-PowerShell" -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "✓ Active Directory module available" -ForegroundColor Green
    }

    # Copy files to installation directory
    Write-Host "Installing files to: $InstallPath" -ForegroundColor Yellow

    $currentDir = $PSScriptRoot
    if (-not $currentDir) {
        $currentDir = Get-Location
    }

    $filesToCopy = @(
        "Add-ForeignUsertoFileShareADGroup.ps1",
        "Add-ForeignUsertoFileShareADGroupGUIController.ps1",
        "Sample_UserGroups.csv",
        "README.md"
    )

    foreach ($file in $filesToCopy) {
        $sourcePath = Join-Path $currentDir $file
        $destPath = Join-Path $InstallPath $file

        if (Test-Path $sourcePath) {
            Copy-Item -Path $sourcePath -Destination $destPath -Force
            Write-Host "✓ Copied: $file" -ForegroundColor Green
        }
        else {
            Write-Warning "File not found: $file"
        }
    }

    # Set execution policy for current user
    Write-Host "Setting execution policy..." -ForegroundColor Yellow
    try {
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
        Write-Host "✓ Execution policy set to RemoteSigned for current user" -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to set execution policy: $($_.Exception.Message)"
    }

    # Create desktop shortcut
    if ($CreateDesktopShortcut) {
        Write-Host "Creating desktop shortcut..." -ForegroundColor Yellow
        try {
            $desktopPath = [Environment]::GetFolderPath("Desktop")
            $shortcutPath = Join-Path $desktopPath "AD Group Management Tool.lnk"
            $guiScriptPath = Join-Path $InstallPath "Add-ForeignUsertoFileShareADGroupGUIController.ps1"

            $shell = New-Object -ComObject WScript.Shell
            $shortcut = $shell.CreateShortcut($shortcutPath)
            $shortcut.TargetPath = "powershell.exe"
            $shortcut.Arguments = "-ExecutionPolicy Bypass -File `"$guiScriptPath`""
            $shortcut.WorkingDirectory = $InstallPath
            $shortcut.Description = "Enhanced AD Group Management Tool"
            $shortcut.Save()

            Write-Host "✓ Desktop shortcut created" -ForegroundColor Green
        }
        catch {
            Write-Warning "Failed to create desktop shortcut: $($_.Exception.Message)"
        }
    }

    # Create sample configuration
    Write-Host "Creating sample configuration..." -ForegroundColor Yellow
    $sampleConfig = @{
        CSVPath            = Join-Path $InstallPath "Sample_UserGroups.csv"
        OutputFolderPath   = Join-Path $InstallPath "Output"
        ScriptPath         = Join-Path $InstallPath "Add-ForeignUsertoFileShareADGroup.ps1"
        EmailFrom          = "noreply@company.com"
        EmailTo            = "admin@company.com"
        EmailSubject       = "AD Group Script Completion"
        EmailBody          = "The Add-ForeignUsertoFileShareADGroup script has completed successfully."
        SMTPServer         = "smtp.company.com"
        SMTPPort           = "587"
        MaxRetries         = 3
        RetryDelaySeconds  = 5
        ParallelProcessing = $false
    }

    # Create output directory
    $outputDir = Join-Path $InstallPath "Output"
    if (-not (Test-Path $outputDir)) {
        New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
    }

    Write-Host ""
    Write-Host "=== Installation Complete ===" -ForegroundColor Green
    Write-Host "Installation Path: $InstallPath" -ForegroundColor Cyan
    Write-Host "To launch the GUI: .\Add-ForeignUsertoFileShareADGroupGUIController.ps1" -ForegroundColor Cyan
    Write-Host "To run via command line: .\Add-ForeignUsertoFileShareADGroup.ps1 -InputCsvPath 'file.csv' -OutputFolderPath 'output'" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Next Steps:" -ForegroundColor Yellow
    Write-Host "1. Review the README.md for detailed usage instructions" -ForegroundColor White
    Write-Host "2. Test with the sample CSV file in test mode" -ForegroundColor White
    Write-Host "3. Configure your actual CSV file and settings" -ForegroundColor White
    Write-Host "4. Ensure you have appropriate domain permissions" -ForegroundColor White
    Write-Host ""

    if (-not $adModule) {
        Write-Host "⚠ IMPORTANT: Install RSAT Tools before using the tool" -ForegroundColor Red
    }

}
catch {
    Write-Error "Installation failed: $($_.Exception.Message)"
    Write-Host "Please check the error and try again, or install manually." -ForegroundColor Red
    exit 1
}

Write-Host "Press any key to continue..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
