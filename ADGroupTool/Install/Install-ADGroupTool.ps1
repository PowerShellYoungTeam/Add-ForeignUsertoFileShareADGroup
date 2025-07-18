function Install-ADGroupTool {
    <#
    .SYNOPSIS
    Installs the AD Group Management Tool module and its dependencies.

    .DESCRIPTION
    This function installs the AD Group Management Tool as a PowerShell module,
    including all dependencies, creating shortcuts, and setting up the environment.

    .PARAMETER InstallPath
    The path where the module will be installed. Default is the user's PowerShell modules directory.

    .PARAMETER CreateDesktopShortcut
    Creates a desktop shortcut to launch the GUI.

    .PARAMETER InstallRSAT
    Attempts to install RSAT tools (requires administrator privileges).

    .PARAMETER Scope
    Installation scope: CurrentUser or AllUsers. Default is CurrentUser.

    .EXAMPLE
    Install-ADGroupTool -CreateDesktopShortcut

    .EXAMPLE
    Install-ADGroupTool -InstallRSAT -Scope AllUsers

    .NOTES
    Author: Steven Wight with GitHub Copilot
    Compatible with: Windows Server 2012 R2+ and PowerShell 5.1+
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$InstallPath,

        [Parameter(Mandatory = $false)]
        [switch]$CreateDesktopShortcut,

        [Parameter(Mandatory = $false)]
        [switch]$InstallRSAT,

        [Parameter(Mandatory = $false)]
        [ValidateSet("CurrentUser", "AllUsers")]
        [string]$Scope = "CurrentUser"
    )

    Write-Host "=== Enhanced AD Group Management Tool - Installation ===" -ForegroundColor Green
    Write-Host ""

    # Determine installation path
    if (-not $InstallPath) {
        if ($Scope -eq "AllUsers") {
            $InstallPath = Join-Path $env:ProgramFiles "WindowsPowerShell\Modules\ADGroupTool"
        } else {
            $userModulesPath = Join-Path ([Environment]::GetFolderPath("MyDocuments")) "WindowsPowerShell\Modules"
            $InstallPath = Join-Path $userModulesPath "ADGroupTool"
        }
    }

    Write-Host "Installation path: $InstallPath" -ForegroundColor Cyan
    Write-Host "Installation scope: $Scope" -ForegroundColor Cyan

    # Check if running as Administrator (required for AllUsers scope or RSAT)
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (($Scope -eq "AllUsers" -or $InstallRSAT) -and -not $isAdmin) {
        Write-Warning "Administrator privileges required for AllUsers installation or RSAT installation."
        if ($Scope -eq "AllUsers") {
            Write-Host "Switching to CurrentUser scope..." -ForegroundColor Yellow
            $Scope = "CurrentUser"
            $userModulesPath = Join-Path ([Environment]::GetFolderPath("MyDocuments")) "WindowsPowerShell\Modules"
            $InstallPath = Join-Path $userModulesPath "ADGroupTool"
        }
        if ($InstallRSAT) {
            Write-Host "Skipping RSAT installation..." -ForegroundColor Yellow
            $InstallRSAT = $false
        }
    }

    try {
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
            Write-Warning "Active Directory PowerShell module is not available."
            if ($InstallRSAT) {
                Write-Host "Attempting to install RSAT tools..." -ForegroundColor Yellow
                try {
                    # Try different methods based on OS version
                    $osVersion = [System.Environment]::OSVersion.Version
                    if ($osVersion.Major -ge 10) {
                        # Windows 10/Server 2016 and later
                        Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -ErrorAction Stop
                    } else {
                        # Windows 8.1/Server 2012 R2 and earlier
                        Enable-WindowsOptionalFeature -Online -FeatureName RSATTools-AD-PowerShell -All -ErrorAction Stop
                    }
                    Write-Host "✓ RSAT tools installed successfully" -ForegroundColor Green
                } catch {
                    Write-Warning "Failed to install RSAT tools automatically: $($_.Exception.Message)"
                    Write-Host "Please install RSAT tools manually:" -ForegroundColor Yellow
                    Write-Host "  Windows 10/2016+: Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'" -ForegroundColor Cyan
                    Write-Host "  Windows 8.1/2012R2: Enable-WindowsOptionalFeature -Online -FeatureName RSATTools-AD-PowerShell" -ForegroundColor Cyan
                }
            } else {
                Write-Host "Please install RSAT tools manually or run with -InstallRSAT parameter" -ForegroundColor Yellow
            }
        } else {
            Write-Host "✓ Active Directory module available: $($adModule.Version)" -ForegroundColor Green
        }

        # Create installation directory structure
        Write-Host "Creating installation directory structure..." -ForegroundColor Yellow

        $directories = @(
            $InstallPath,
            (Join-Path $InstallPath "Functions"),
            (Join-Path $InstallPath "GUI"),
            (Join-Path $InstallPath "Install"),
            (Join-Path $InstallPath "Data")
        )

        foreach ($dir in $directories) {
            if (-not (Test-Path $dir)) {
                New-Item -Path $dir -ItemType Directory -Force | Out-Null
                Write-Host "  Created: $dir" -ForegroundColor Gray
            }
        }

        # Get source files location
        $sourceRoot = $PSScriptRoot
        if (-not $sourceRoot) {
            # Fallback to current directory
            $sourceRoot = (Get-Location).Path
        }

        # Find the ADGroupTool directory
        $moduleSourcePath = $null
        $possiblePaths = @(
            (Join-Path $sourceRoot "ADGroupTool"),
            (Join-Path (Split-Path $sourceRoot -Parent) "ADGroupTool"),
            $sourceRoot
        )

        foreach ($path in $possiblePaths) {
            if (Test-Path (Join-Path $path "ADGroupTool.psd1")) {
                $moduleSourcePath = $path
                break
            }
        }

        if (-not $moduleSourcePath) {
            throw "Could not locate ADGroupTool module source files. Please ensure you're running this from the correct directory."
        }

        Write-Host "Source location: $moduleSourcePath" -ForegroundColor Cyan

        # Copy module files
        Write-Host "Installing module files..." -ForegroundColor Yellow

        $filesToCopy = @(
            @{ Source = "ADGroupTool.psd1"; Destination = "" },
            @{ Source = "ADGroupTool.psm1"; Destination = "" },
            @{ Source = "Functions\*.ps1"; Destination = "Functions" },
            @{ Source = "GUI\*.ps1"; Destination = "GUI" },
            @{ Source = "Install\*.ps1"; Destination = "Install" }
        )

        # Add sample data if available
        $sampleDataPath = Join-Path (Split-Path $moduleSourcePath -Parent) "Sample_UserGroups.csv"
        if (Test-Path $sampleDataPath) {
            $filesToCopy += @{ Source = "..\Sample_UserGroups.csv"; Destination = "Data\Sample_UserGroups.csv" }
        }

        foreach ($file in $filesToCopy) {
            $sourcePath = Join-Path $moduleSourcePath $file.Source
            $destinationPath = if ($file.Destination) {
                Join-Path $InstallPath $file.Destination
            } else {
                $InstallPath
            }

            if ($file.Source -like "*\*") {
                # Handle wildcards
                $sourceFiles = Get-ChildItem -Path $sourcePath -File
                foreach ($sourceFile in $sourceFiles) {
                    $destFile = Join-Path $destinationPath $sourceFile.Name
                    Copy-Item -Path $sourceFile.FullName -Destination $destFile -Force
                    Write-Host "  Copied: $($sourceFile.Name)" -ForegroundColor Gray
                }
            } else {
                if (Test-Path $sourcePath) {
                    if ($file.Destination -and $file.Destination.Contains("\")) {
                        # Specific file destination
                        $destFile = Join-Path $InstallPath $file.Destination
                    } else {
                        # Copy to folder
                        $destFile = Join-Path $destinationPath (Split-Path $sourcePath -Leaf)
                    }
                    Copy-Item -Path $sourcePath -Destination $destFile -Force
                    Write-Host "  Copied: $(Split-Path $sourcePath -Leaf)" -ForegroundColor Gray
                }
            }
        }

        # Import the module to test installation
        Write-Host "Testing module installation..." -ForegroundColor Yellow
        try {
            Import-Module $InstallPath -Force -ErrorAction Stop
            $module = Get-Module -Name ADGroupTool
            Write-Host "✓ Module imported successfully - Version: $($module.Version)" -ForegroundColor Green

            # Test key functions
            $functions = @("Test-CSVSchema", "Test-Credential", "Test-DomainConnectivity", "Add-ADUserToGroup", "Invoke-ADGroupOperation", "Start-ADGroupToolGUI")
            $availableFunctions = Get-Command -Module ADGroupTool | Select-Object -ExpandProperty Name

            foreach ($func in $functions) {
                if ($func -in $availableFunctions) {
                    Write-Host "  ✓ Function available: $func" -ForegroundColor Gray
                } else {
                    Write-Warning "  ✗ Function missing: $func"
                }
            }
        } catch {
            Write-Warning "Module import test failed: $($_.Exception.Message)"
        }

        # Create desktop shortcut if requested
        if ($CreateDesktopShortcut) {
            Write-Host "Creating desktop shortcut..." -ForegroundColor Yellow
            try {
                $desktop = [Environment]::GetFolderPath("Desktop")
                $shortcutPath = Join-Path $desktop "AD Group Tool.lnk"

                $WScript = New-Object -ComObject WScript.Shell
                $shortcut = $WScript.CreateShortcut($shortcutPath)
                $shortcut.TargetPath = "powershell.exe"
                $shortcut.Arguments = "-Command `"Import-Module ADGroupTool; Start-ADGroupToolGUI`""
                $shortcut.WorkingDirectory = $InstallPath
                $shortcut.Description = "Enhanced AD Group Management Tool"
                $shortcut.Save()

                Write-Host "✓ Desktop shortcut created: $shortcutPath" -ForegroundColor Green
            } catch {
                Write-Warning "Failed to create desktop shortcut: $($_.Exception.Message)"
            }
        }

        # Add to PowerShell profile (optional)
        Write-Host "Updating PowerShell profile..." -ForegroundColor Yellow
        try {
            $profilePath = $PROFILE.CurrentUserAllHosts
            $profileDir = Split-Path $profilePath -Parent

            if (-not (Test-Path $profileDir)) {
                New-Item -Path $profileDir -ItemType Directory -Force | Out-Null
            }

            $importCommand = "Import-Module ADGroupTool -DisableNameChecking"

            if (Test-Path $profilePath) {
                $profileContent = Get-Content $profilePath -Raw
                if ($profileContent -notlike "*Import-Module ADGroupTool*") {
                    Add-Content -Path $profilePath -Value "`n# AD Group Management Tool"
                    Add-Content -Path $profilePath -Value $importCommand
                    Write-Host "✓ Added module import to PowerShell profile" -ForegroundColor Green
                } else {
                    Write-Host "✓ Module import already exists in PowerShell profile" -ForegroundColor Green
                }
            } else {
                "# AD Group Management Tool" | Out-File -FilePath $profilePath -Encoding UTF8
                $importCommand | Add-Content -Path $profilePath
                Write-Host "✓ Created PowerShell profile with module import" -ForegroundColor Green
            }
        } catch {
            Write-Warning "Failed to update PowerShell profile: $($_.Exception.Message)"
        }

        # Installation summary
        Write-Host ""
        Write-Host "=== Installation Complete ===" -ForegroundColor Green
        Write-Host "Module installed to: $InstallPath" -ForegroundColor Cyan
        Write-Host "Installation scope: $Scope" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Usage:" -ForegroundColor Yellow
        Write-Host "  Start-ADGroupToolGUI                 # Launch GUI interface" -ForegroundColor Cyan
        Write-Host "  Invoke-ADGroupOperation               # Command-line interface" -ForegroundColor Cyan
        Write-Host "  Get-Help Invoke-ADGroupOperation      # Get detailed help" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Examples:" -ForegroundColor Yellow
        Write-Host "  Import-Module ADGroupTool" -ForegroundColor Cyan
        Write-Host "  Start-ADGroupToolGUI" -ForegroundColor Cyan
        Write-Host ""

        if ($CreateDesktopShortcut) {
            Write-Host "Desktop shortcut created for easy access!" -ForegroundColor Green
        }

        return $true
    }
    catch {
        Write-Error "Installation failed: $($_.Exception.Message)"
        Write-Host "Check the error details above and try again." -ForegroundColor Red
        return $false
    }
}
