function Start-ADGroupToolGUI {
    <#
    .SYNOPSIS
    Launches the enhanced GUI interface for the AD Group Management Tool.

    .DESCRIPTION
    This function provides a modern tabbed GUI interface for configuring and executing
    AD group operations. Features include real-time validation, configuration persistence,
    progress monitoring, and comprehensive logging.

    .PARAMETER ConfigPath
    Optional path to a custom configuration file. If not specified, uses the default user config location.

    .EXAMPLE
    Start-ADGroupToolGUI

    .NOTES
    Author: Steven Wight with GitHub Copilot
    Compatible with: Windows Server 2012 R2+ and PowerShell 5.1+
    Requires: System.Windows.Forms, System.Drawing
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$ConfigPath
    )

    # Load required assemblies
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Global variables for better state management
    $script:foreignAdminCreds = $null
    $script:configFile = if ($ConfigPath) { $ConfigPath } else { Join-Path $env:APPDATA "ADGroupScriptGUI_Config.xml" }
    $script:logTextBox = $null
    $script:isRunning = $false

    # Default configuration
    $defaultConfig = @{
        CSVPath             = ""
        OutputFolderPath    = "C:\temp\PowerShell"
        EmailFrom           = "noreply@company.com"
        EmailTo             = "admin@company.com"
        EmailSubject        = "AD Group Script Completion"
        EmailBody           = "The AD Group Management Tool has completed successfully. Please see the attached output file for details."
        SMTPServer          = "smtp.company.com"
        SMTPPort            = "587"
        MaxRetries          = 3
        RetryDelaySeconds   = 5
        RememberCredentials = $false
        ValidateCredentials = $true
        TestConnectivity    = $true
        ExponentialBackoff  = $false
    }

    # Enhanced Functions for better functionality
    function Save-Configuration {
        param([hashtable]$Config)
        try {
            $Config | Export-Clixml -Path $script:configFile -Force
            Write-Log "Configuration saved successfully" "Info"
            return $true
        }
        catch {
            Write-Log "Failed to save configuration: $($_.Exception.Message)" "Error"
            return $false
        }
    }

    function Load-Configuration {
        try {
            if (Test-Path $script:configFile) {
                $config = Import-Clixml -Path $script:configFile
                Write-Log "Configuration loaded successfully" "Info"
                return $config
            }
        }
        catch {
            Write-Log "Failed to load configuration: $($_.Exception.Message)" "Warning"
        }
        return $defaultConfig
    }

    function Write-Log {
        param(
            [string]$Message,
            [ValidateSet("Info", "Warning", "Error", "Success")]
            [string]$Level = "Info"
        )

        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logMessage = "[$timestamp] [$Level] $Message"

        if ($script:logTextBox) {
            $script:logTextBox.AppendText("$logMessage`r`n")
            $script:logTextBox.ScrollToCaret()
        }

        # Also write to console
        Write-Verbose $logMessage
    }

    function Test-Prerequisites {
        $issues = @()

        # Check PowerShell version
        if ($PSVersionTable.PSVersion.Major -lt 5) {
            $issues += "PowerShell 5.1 or later is required. Current version: $($PSVersionTable.PSVersion)"
        }

        # Check for Active Directory module
        if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
            $issues += "Active Directory PowerShell module is not available. Please install RSAT tools."
        }

        # Check platform
        if ($PSVersionTable.Platform -and $PSVersionTable.Platform -ne "Win32NT") {
            $issues += "This tool requires Windows platform."
        }

        return $issues
    }

    function Validate-CSVFile {
        param([string]$CsvPath)

        if (-not $CsvPath -or -not (Test-Path $CsvPath)) {
            return @{ IsValid = $false; Message = "CSV file not found or path is empty" }
        }

        try {
            $validation = Test-CSVSchema -CsvPath $CsvPath
            if ($validation.IsValid) {
                return @{
                    IsValid = $true;
                    Message = "Valid CSV: $($validation.ValidRows)/$($validation.TotalRows) valid rows"
                }
            }
            else {
                return @{
                    IsValid = $false;
                    Message = "CSV validation failed: $($validation.Error)"
                }
            }
        }
        catch {
            return @{ IsValid = $false; Message = "Error validating CSV: $($_.Exception.Message)" }
        }
    }

    function Show-CSVPreview {
        param([string]$CsvPath)

        if (-not (Test-Path $CsvPath)) {
            [System.Windows.Forms.MessageBox]::Show("CSV file not found: $CsvPath", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        try {
            $csvData = Import-Csv -Path $CsvPath | Select-Object -First 10
            $preview = $csvData | Format-Table -AutoSize | Out-String

            $previewForm = New-Object System.Windows.Forms.Form
            $previewForm.Text = "CSV Preview (First 10 Rows)"
            $previewForm.Size = New-Object System.Drawing.Size(800, 600)
            $previewForm.StartPosition = "CenterParent"

            $textBox = New-Object System.Windows.Forms.TextBox
            $textBox.Multiline = $true
            $textBox.ReadOnly = $true
            $textBox.ScrollBars = "Both"
            $textBox.Font = New-Object System.Drawing.Font("Consolas", 9)
            $textBox.Dock = "Fill"
            $textBox.Text = $preview

            $previewForm.Controls.Add($textBox)
            $previewForm.ShowDialog()
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error reading CSV file: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }

    # Check prerequisites before starting GUI
    $prerequisiteIssues = Test-Prerequisites
    if ($prerequisiteIssues.Count -gt 0) {
        $message = "Prerequisites check failed:`n`n" + ($prerequisiteIssues -join "`n")
        [System.Windows.Forms.MessageBox]::Show($message, "Prerequisites Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }

    # Load configuration
    $config = Load-Configuration

    # Create main form
    $mainForm = New-Object System.Windows.Forms.Form
    $mainForm.Text = "Enhanced AD Group Management Tool v2.0"
    $mainForm.Size = New-Object System.Drawing.Size(900, 700)
    $mainForm.StartPosition = "CenterScreen"
    $mainForm.FormBorderStyle = "FixedDialog"
    $mainForm.MaximizeBox = $false

    # Create tab control
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = "Fill"
    $tabControl.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # Configuration Tab
    $configTab = New-Object System.Windows.Forms.TabPage
    $configTab.Text = "Configuration"
    $configTab.Padding = New-Object System.Windows.Forms.Padding(10)

    # CSV File Section
    $csvLabel = New-Object System.Windows.Forms.Label
    $csvLabel.Text = "CSV File Path:"
    $csvLabel.Location = New-Object System.Drawing.Point(10, 20)
    $csvLabel.Size = New-Object System.Drawing.Size(100, 20)

    $csvTextBox = New-Object System.Windows.Forms.TextBox
    $csvTextBox.Location = New-Object System.Drawing.Point(120, 18)
    $csvTextBox.Size = New-Object System.Drawing.Size(500, 22)
    $csvTextBox.Text = $config.CSVPath

    $csvBrowseButton = New-Object System.Windows.Forms.Button
    $csvBrowseButton.Text = "Browse"
    $csvBrowseButton.Location = New-Object System.Drawing.Point(630, 17)
    $csvBrowseButton.Size = New-Object System.Drawing.Size(70, 25)

    $csvPreviewButton = New-Object System.Windows.Forms.Button
    $csvPreviewButton.Text = "Preview"
    $csvPreviewButton.Location = New-Object System.Drawing.Point(710, 17)
    $csvPreviewButton.Size = New-Object System.Drawing.Size(70, 25)

    $csvValidationLabel = New-Object System.Windows.Forms.Label
    $csvValidationLabel.Location = New-Object System.Drawing.Point(120, 45)
    $csvValidationLabel.Size = New-Object System.Drawing.Size(660, 20)
    $csvValidationLabel.ForeColor = [System.Drawing.Color]::Gray
    $csvValidationLabel.Text = "Select a CSV file to validate"

    # Output Folder Section
    $outputLabel = New-Object System.Windows.Forms.Label
    $outputLabel.Text = "Output Folder:"
    $outputLabel.Location = New-Object System.Drawing.Point(10, 80)
    $outputLabel.Size = New-Object System.Drawing.Size(100, 20)

    $outputTextBox = New-Object System.Windows.Forms.TextBox
    $outputTextBox.Location = New-Object System.Drawing.Point(120, 78)
    $outputTextBox.Size = New-Object System.Drawing.Size(500, 22)
    $outputTextBox.Text = $config.OutputFolderPath

    $outputBrowseButton = New-Object System.Windows.Forms.Button
    $outputBrowseButton.Text = "Browse"
    $outputBrowseButton.Location = New-Object System.Drawing.Point(630, 77)
    $outputBrowseButton.Size = New-Object System.Drawing.Size(70, 25)

    # Mode Selection
    $modeGroupBox = New-Object System.Windows.Forms.GroupBox
    $modeGroupBox.Text = "Execution Mode"
    $modeGroupBox.Location = New-Object System.Drawing.Point(10, 120)
    $modeGroupBox.Size = New-Object System.Drawing.Size(780, 60)

    $testModeRadio = New-Object System.Windows.Forms.RadioButton
    $testModeRadio.Text = "Test Mode (WhatIf - No changes made)"
    $testModeRadio.Location = New-Object System.Drawing.Point(20, 25)
    $testModeRadio.Size = New-Object System.Drawing.Size(300, 20)
    $testModeRadio.Checked = $true

    $liveModeRadio = New-Object System.Windows.Forms.RadioButton
    $liveModeRadio.Text = "Live Mode (Apply changes)"
    $liveModeRadio.Location = New-Object System.Drawing.Point(400, 25)
    $liveModeRadio.Size = New-Object System.Drawing.Size(200, 20)

    # Credentials Section
    $credGroupBox = New-Object System.Windows.Forms.GroupBox
    $credGroupBox.Text = "Domain Credentials"
    $credGroupBox.Location = New-Object System.Drawing.Point(10, 200)
    $credGroupBox.Size = New-Object System.Drawing.Size(780, 80)

    $credButton = New-Object System.Windows.Forms.Button
    $credButton.Text = "Set Credentials"
    $credButton.Location = New-Object System.Drawing.Point(20, 30)
    $credButton.Size = New-Object System.Drawing.Size(120, 30)

    $credStatusLabel = New-Object System.Windows.Forms.Label
    $credStatusLabel.Text = "No credentials set"
    $credStatusLabel.Location = New-Object System.Drawing.Point(160, 35)
    $credStatusLabel.Size = New-Object System.Drawing.Size(400, 20)
    $credStatusLabel.ForeColor = [System.Drawing.Color]::Red

    # Run Button
    $runButton = New-Object System.Windows.Forms.Button
    $runButton.Text = "Run Operation"
    $runButton.Location = New-Object System.Drawing.Point(350, 300)
    $runButton.Size = New-Object System.Drawing.Size(120, 40)
    $runButton.BackColor = [System.Drawing.Color]::LightGreen
    $runButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

    # Advanced Tab
    $advancedTab = New-Object System.Windows.Forms.TabPage
    $advancedTab.Text = "Advanced"
    $advancedTab.Padding = New-Object System.Windows.Forms.Padding(10)

    # Retry Settings
    $retryGroupBox = New-Object System.Windows.Forms.GroupBox
    $retryGroupBox.Text = "Retry Settings"
    $retryGroupBox.Location = New-Object System.Drawing.Point(10, 20)
    $retryGroupBox.Size = New-Object System.Drawing.Size(780, 100)

    $maxRetriesLabel = New-Object System.Windows.Forms.Label
    $maxRetriesLabel.Text = "Max Retries:"
    $maxRetriesLabel.Location = New-Object System.Drawing.Point(20, 30)
    $maxRetriesLabel.Size = New-Object System.Drawing.Size(80, 20)

    $maxRetriesNumeric = New-Object System.Windows.Forms.NumericUpDown
    $maxRetriesNumeric.Location = New-Object System.Drawing.Point(110, 28)
    $maxRetriesNumeric.Size = New-Object System.Drawing.Size(60, 22)
    $maxRetriesNumeric.Minimum = 1
    $maxRetriesNumeric.Maximum = 10
    $maxRetriesNumeric.Value = $config.MaxRetries

    $retryDelayLabel = New-Object System.Windows.Forms.Label
    $retryDelayLabel.Text = "Retry Delay (sec):"
    $retryDelayLabel.Location = New-Object System.Drawing.Point(200, 30)
    $retryDelayLabel.Size = New-Object System.Drawing.Size(100, 20)

    $retryDelayNumeric = New-Object System.Windows.Forms.NumericUpDown
    $retryDelayNumeric.Location = New-Object System.Drawing.Point(310, 28)
    $retryDelayNumeric.Size = New-Object System.Drawing.Size(60, 22)
    $retryDelayNumeric.Minimum = 1
    $retryDelayNumeric.Maximum = 60
    $retryDelayNumeric.Value = $config.RetryDelaySeconds

    $exponentialBackoffCheckBox = New-Object System.Windows.Forms.CheckBox
    $exponentialBackoffCheckBox.Text = "Use Exponential Backoff"
    $exponentialBackoffCheckBox.Location = New-Object System.Drawing.Point(20, 60)
    $exponentialBackoffCheckBox.Size = New-Object System.Drawing.Size(200, 20)
    $exponentialBackoffCheckBox.Checked = $config.ExponentialBackoff

    # Validation Settings
    $validationGroupBox = New-Object System.Windows.Forms.GroupBox
    $validationGroupBox.Text = "Pre-flight Validation"
    $validationGroupBox.Location = New-Object System.Drawing.Point(10, 140)
    $validationGroupBox.Size = New-Object System.Drawing.Size(780, 80)

    $validateCredsCheckBox = New-Object System.Windows.Forms.CheckBox
    $validateCredsCheckBox.Text = "Validate Credentials"
    $validateCredsCheckBox.Location = New-Object System.Drawing.Point(20, 30)
    $validateCredsCheckBox.Size = New-Object System.Drawing.Size(150, 20)
    $validateCredsCheckBox.Checked = $config.ValidateCredentials

    $testConnectivityCheckBox = New-Object System.Windows.Forms.CheckBox
    $testConnectivityCheckBox.Text = "Test Domain Connectivity"
    $testConnectivityCheckBox.Location = New-Object System.Drawing.Point(200, 30)
    $testConnectivityCheckBox.Size = New-Object System.Drawing.Size(180, 20)
    $testConnectivityCheckBox.Checked = $config.TestConnectivity

    # Logs Tab
    $logsTab = New-Object System.Windows.Forms.TabPage
    $logsTab.Text = "Logs"
    $logsTab.Padding = New-Object System.Windows.Forms.Padding(10)

    $script:logTextBox = New-Object System.Windows.Forms.TextBox
    $script:logTextBox.Multiline = $true
    $script:logTextBox.ReadOnly = $true
    $script:logTextBox.ScrollBars = "Both"
    $script:logTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $script:logTextBox.Dock = "Fill"

    $clearLogsButton = New-Object System.Windows.Forms.Button
    $clearLogsButton.Text = "Clear Logs"
    $clearLogsButton.Location = New-Object System.Drawing.Point(10, 10)
    $clearLogsButton.Size = New-Object System.Drawing.Size(80, 25)
    $clearLogsButton.Dock = "Bottom"

    # Event Handlers
    $csvBrowseButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $openFileDialog.Title = "Select CSV File"

        if ($openFileDialog.ShowDialog() -eq "OK") {
            $csvTextBox.Text = $openFileDialog.FileName
            # Trigger validation
            $csvTextBox_TextChanged.Invoke()
        }
    })

    $csvPreviewButton.Add_Click({
        if ($csvTextBox.Text) {
            Show-CSVPreview -CsvPath $csvTextBox.Text
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Please select a CSV file first.", "No File Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })

    $csvTextBox_TextChanged = {
        if ($csvTextBox.Text) {
            $validation = Validate-CSVFile -CsvPath $csvTextBox.Text
            if ($validation.IsValid) {
                $csvValidationLabel.Text = $validation.Message
                $csvValidationLabel.ForeColor = [System.Drawing.Color]::Green
            }
            else {
                $csvValidationLabel.Text = $validation.Message
                $csvValidationLabel.ForeColor = [System.Drawing.Color]::Red
            }
        }
        else {
            $csvValidationLabel.Text = "Select a CSV file to validate"
            $csvValidationLabel.ForeColor = [System.Drawing.Color]::Gray
        }
    }
    $csvTextBox.Add_TextChanged($csvTextBox_TextChanged)

    $outputBrowseButton.Add_Click({
        $folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowserDialog.Description = "Select Output Folder"
        $folderBrowserDialog.SelectedPath = $outputTextBox.Text

        if ($folderBrowserDialog.ShowDialog() -eq "OK") {
            $outputTextBox.Text = $folderBrowserDialog.SelectedPath
        }
    })

    $credButton.Add_Click({
        $script:foreignAdminCreds = Get-Credential -Message "Enter foreign domain admin credentials"
        if ($script:foreignAdminCreds) {
            $credStatusLabel.Text = "Credentials set for: $($script:foreignAdminCreds.UserName)"
            $credStatusLabel.ForeColor = [System.Drawing.Color]::Green
        }
    })

    $runButton.Add_Click({
        if ($script:isRunning) {
            [System.Windows.Forms.MessageBox]::Show("Operation is already running. Please wait for it to complete.", "Operation in Progress", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }

        # Validate inputs
        if (-not $csvTextBox.Text -or -not (Test-Path $csvTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please select a valid CSV file.", "Invalid Input", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        if (-not $outputTextBox.Text) {
            [System.Windows.Forms.MessageBox]::Show("Please select an output folder.", "Invalid Input", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        if (-not $script:foreignAdminCreds) {
            [System.Windows.Forms.MessageBox]::Show("Please set domain credentials.", "Missing Credentials", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        # Save configuration
        $currentConfig = @{
            CSVPath             = $csvTextBox.Text
            OutputFolderPath    = $outputTextBox.Text
            MaxRetries          = $maxRetriesNumeric.Value
            RetryDelaySeconds   = $retryDelayNumeric.Value
            ExponentialBackoff  = $exponentialBackoffCheckBox.Checked
            ValidateCredentials = $validateCredsCheckBox.Checked
            TestConnectivity    = $testConnectivityCheckBox.Checked
        }
        Save-Configuration -Config $currentConfig

        # Switch to logs tab
        $tabControl.SelectedTab = $logsTab

        $script:isRunning = $true
        $runButton.Enabled = $false
        $runButton.Text = "Running..."

        try {
            Write-Log "Starting AD Group operation..." "Info"
            Write-Log "CSV File: $($csvTextBox.Text)" "Info"
            Write-Log "Output Folder: $($outputTextBox.Text)" "Info"
            Write-Log "Mode: $(if ($testModeRadio.Checked) { 'Test' } else { 'Live' })" "Info"

            # Execute the operation
            $result = Invoke-ADGroupOperation -InputCsvPath $csvTextBox.Text -OutputFolderPath $outputTextBox.Text -Test:$testModeRadio.Checked -ForeignAdminCreds $script:foreignAdminCreds -MaxRetries $maxRetriesNumeric.Value -RetryDelaySeconds $retryDelayNumeric.Value -ExponentialBackoff:$exponentialBackoffCheckBox.Checked -ValidateCredentials:$validateCredsCheckBox.Checked -TestConnectivity:$testConnectivityCheckBox.Checked -Verbose

            if ($result) {
                Write-Log "Operation completed successfully!" "Success"
            }
            else {
                Write-Log "Operation completed with errors. Check the output folder for details." "Warning"
            }
        }
        catch {
            Write-Log "Operation failed: $($_.Exception.Message)" "Error"
        }
        finally {
            $script:isRunning = $false
            $runButton.Enabled = $true
            $runButton.Text = "Run Operation"
        }
    })

    $clearLogsButton.Add_Click({
        $script:logTextBox.Clear()
    })

    # Add controls to tabs
    $configTab.Controls.AddRange(@($csvLabel, $csvTextBox, $csvBrowseButton, $csvPreviewButton, $csvValidationLabel, $outputLabel, $outputTextBox, $outputBrowseButton, $modeGroupBox, $credGroupBox, $runButton))
    $modeGroupBox.Controls.AddRange(@($testModeRadio, $liveModeRadio))
    $credGroupBox.Controls.AddRange(@($credButton, $credStatusLabel))

    $advancedTab.Controls.AddRange(@($retryGroupBox, $validationGroupBox))
    $retryGroupBox.Controls.AddRange(@($maxRetriesLabel, $maxRetriesNumeric, $retryDelayLabel, $retryDelayNumeric, $exponentialBackoffCheckBox))
    $validationGroupBox.Controls.AddRange(@($validateCredsCheckBox, $testConnectivityCheckBox))

    $logsTab.Controls.AddRange(@($script:logTextBox, $clearLogsButton))

    # Add tabs to control
    $tabControl.TabPages.Add($configTab)
    $tabControl.TabPages.Add($advancedTab)
    $tabControl.TabPages.Add($logsTab)

    # Add tab control to form
    $mainForm.Controls.Add($tabControl)

    # Initialize form
    Write-Log "AD Group Management Tool GUI initialized" "Info"
    Write-Log "Configuration loaded from: $script:configFile" "Info"

    # Trigger initial validation if CSV path is set
    if ($csvTextBox.Text) {
        $csvTextBox_TextChanged.Invoke()
    }

    # Show form
    $mainForm.ShowDialog()
}
