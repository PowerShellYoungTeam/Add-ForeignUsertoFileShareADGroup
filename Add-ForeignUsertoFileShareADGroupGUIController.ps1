# Enhanced GUI Controller for Add-ForeignUsertoFileShareADGroup Script
# Author: Steven Wight with GitHub Copilot
# Compatible with Windows PowerShell 5.1
# Enhanced with comprehensive validation, better UX, and modern features

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global variables for better state management
$script:foreignAdminCreds = $null
$script:configFile = Join-Path $env:APPDATA "ADGroupScriptGUI_Config.xml"
$script:logTextBox = $null
$script:isRunning = $false

# Default configuration
$defaultConfig = @{
    CSVPath             = "C:\temp\Powershell\BarryGroups.csv"
    OutputFolderPath    = "C:\temp\PowerShell"
    EmailFrom           = "noreply@company.com"
    EmailTo             = "admin@company.com"
    EmailSubject        = "AD Group Script Completion"
    EmailBody           = "The Add-ForeignUsertoFileShareADGroup script has completed successfully. Please see the attached output file for details."
    SMTPServer          = "smtp.company.com"
    SMTPPort            = "587"
    ScriptPath          = ""
    MaxRetries          = 3
    RetryDelaySeconds   = 5
    RememberCredentials = $false
}

# Enhanced Functions for better functionality

function Save-Configuration {
    param([hashtable]$Config)
    try {
        $Config | Export-Clixml -Path $script:configFile -Force
        Write-Log "Configuration saved successfully" "Info"
    }
    catch {
        Write-Log "Failed to save configuration: $($_.Exception.Message)" "Error"
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
        # Color coding for different log levels
        $color = switch ($Level) {
            "Info" { [System.Drawing.Color]::Black }
            "Warning" { [System.Drawing.Color]::Orange }
            "Error" { [System.Drawing.Color]::Red }
            "Success" { [System.Drawing.Color]::Green }
        }

        $script:logTextBox.AppendText("$logMessage`r`n")
        $script:logTextBox.ScrollToCaret()
    }

    # Also write to console
    Write-Host $logMessage -ForegroundColor $(
        switch ($Level) {
            "Info" { "White" }
            "Warning" { "Yellow" }
            "Error" { "Red" }
            "Success" { "Green" }
        }
    )
}

function Test-Prerequisites {
    $issues = @()

    # Check PowerShell version
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        $issues += "PowerShell 5.1 or later is required. Current version: $($PSVersionTable.PSVersion)"
    }

    # Check for Active Directory module
    if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
        $issues += "ActiveDirectory PowerShell module is not available. Please install RSAT tools."
    }

    # Check if running on Windows
    if ($PSVersionTable.Platform -and $PSVersionTable.Platform -ne "Win32NT") {
        $issues += "This script requires Windows PowerShell on Windows."
    }

    return $issues
}

function Validate-CSVFile {
    param([string]$CsvPath)

    if (-not (Test-Path $CsvPath)) {
        return @("CSV file does not exist: $CsvPath")
    }

    try {
        $csvData = Import-Csv -Path $CsvPath -ErrorAction Stop
        $requiredColumns = @('SourceDomain', 'SourceUser', 'TargetDomain', 'TargetGroup')

        if ($csvData.Count -eq 0) {
            return @("CSV file is empty")
        }

        $csvColumns = $csvData[0].PSObject.Properties.Name
        $missingColumns = $requiredColumns | Where-Object { $_ -notin $csvColumns }

        $issues = @()
        if ($missingColumns) {
            $issues += "Missing required columns: $($missingColumns -join ', ')"
        }

        # Check for empty data
        $emptyRows = $csvData | Where-Object {
            [string]::IsNullOrWhiteSpace($_.SourceDomain) -or
            [string]::IsNullOrWhiteSpace($_.SourceUser) -or
            [string]::IsNullOrWhiteSpace($_.TargetDomain) -or
            [string]::IsNullOrWhiteSpace($_.TargetGroup)
        }

        if ($emptyRows.Count -eq $csvData.Count) {
            $issues += "All rows contain missing data"
        }
        elseif ($emptyRows.Count -gt 0) {
            $issues += "Warning: $($emptyRows.Count) rows contain missing data and will be skipped"
        }

        return $issues
    }
    catch {
        return @("Failed to read CSV file: $($_.Exception.Message)")
    }
}

function Update-ValidationStatus {
    param(
        [System.Windows.Forms.TextBox]$TextBox,
        [System.Windows.Forms.Label]$StatusLabel,
        [string]$ValidationFunction
    )

    $text = $TextBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($text) -or $text -eq "Choose file" -or $text -eq "Choose folder") {
        $StatusLabel.Text = "Warning"
        $StatusLabel.ForeColor = [System.Drawing.Color]::Orange
        $StatusLabel.BackColor = [System.Drawing.Color]::Transparent
        return $false
    }

    $isValid = $true
    $issues = @()

    switch ($ValidationFunction) {
        "File" {
            if (-not (Test-Path $text -PathType Leaf)) {
                $isValid = $false
                $issues += "File does not exist"
            }
        }
        "Folder" {
            if (-not (Test-Path $text -PathType Container)) {
                $isValid = $false
                $issues += "Folder does not exist"
            }
        }
        "CSV" {
            $issues = Validate-CSVFile -CsvPath $text
            $isValid = $issues.Count -eq 0 -or ($issues.Count -eq 1 -and $issues[0] -like "Warning*")
        }
        "Email" {
            if ($text -notmatch "^[^@]+@[^@]+\.[^@]+$") {
                $isValid = $false
                $issues += "Invalid email format"
            }
        }
        "Number" {
            if (-not ($text -match "^\d+$")) {
                $isValid = $false
                $issues += "Must be a number"
            }
        }
    }

    if ($isValid) {
        $StatusLabel.Text = "OK"
        $StatusLabel.ForeColor = [System.Drawing.Color]::Green
        if ($issues.Count -gt 0 -and $issues[0] -like "Warning*") {
            $StatusLabel.ForeColor = [System.Drawing.Color]::Orange
        }
    }
    else {
        $StatusLabel.Text = "Error"
        $StatusLabel.ForeColor = [System.Drawing.Color]::Red
    }

    $StatusLabel.BackColor = [System.Drawing.Color]::Transparent

    # Set tooltip with validation messages
    if ($issues.Count -gt 0) {
        $tooltip = New-Object System.Windows.Forms.ToolTip
        $tooltip.SetToolTip($StatusLabel, ($issues -join "`n"))
    }

    return $isValid
}
# Load configuration
$config = Load-Configuration

# Check prerequisites
$prereqIssues = Test-Prerequisites
if ($prereqIssues.Count -gt 0) {
    [System.Windows.Forms.MessageBox]::Show(
        "Prerequisites check failed:`n`n$($prereqIssues -join "`n")",
        "Prerequisites Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit 1
}

# Create enhanced main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Enhanced AD Group Management Tool v2.0"
$form.Size = New-Object System.Drawing.Size(800, 750)
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$form.ShowIcon = $true

# Create tab control for better organization
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 50)
$tabControl.Size = New-Object System.Drawing.Size(760, 650)
$form.Controls.Add($tabControl)

# Tab 1: Main Configuration
$mainTab = New-Object System.Windows.Forms.TabPage
$mainTab.Text = "Configuration"
$tabControl.TabPages.Add($mainTab)

# Tab 2: Advanced Options
$advancedTab = New-Object System.Windows.Forms.TabPage
$advancedTab.Text = "Advanced"
$tabControl.TabPages.Add($advancedTab)

# Tab 3: Email Settings
$emailTab = New-Object System.Windows.Forms.TabPage
$emailTab.Text = "Email"
$tabControl.TabPages.Add($emailTab)

# Tab 4: Logs and Results
$logsTab = New-Object System.Windows.Forms.TabPage
$logsTab.Text = "Logs and Results"
$tabControl.TabPages.Add($logsTab)

# Status label at top of form
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Status: Ready"
$statusLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
$statusLabel.Location = New-Object System.Drawing.Point(10, 10)
$statusLabel.Size = New-Object System.Drawing.Size(600, 30)
$statusLabel.ForeColor = [System.Drawing.Color]::Green
$form.Controls.Add($statusLabel)

# === MAIN TAB CONTROLS ===
$yPos = 20

# Script Path Section
$scriptPathLabel = New-Object System.Windows.Forms.Label
$scriptPathLabel.Text = "Script Path:"
$scriptPathLabel.Location = New-Object System.Drawing.Point(10, $yPos)
$scriptPathLabel.Size = New-Object System.Drawing.Size(100, 20)
$mainTab.Controls.Add($scriptPathLabel)

$scriptPathTextBox = New-Object System.Windows.Forms.TextBox
$scriptPathTextBox.Text = $config.ScriptPath
$scriptPathTextBox.Location = New-Object System.Drawing.Point(120, $yPos)
$scriptPathTextBox.Size = New-Object System.Drawing.Size(500, 20)
$mainTab.Controls.Add($scriptPathTextBox)

$scriptPathButton = New-Object System.Windows.Forms.Button
$scriptPathButton.Text = "Browse"
$scriptPathButton.Location = New-Object System.Drawing.Point(630, $yPos)
$scriptPathButton.Size = New-Object System.Drawing.Size(75, 23)
$scriptPathButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "PowerShell files (*.ps1)|*.ps1|All files (*.*)|*.*"
        $openFileDialog.Title = "Select Add-ForeignUsertoFileShareADGroup.ps1"
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $scriptPathTextBox.Text = $openFileDialog.FileName
        }
    })
$mainTab.Controls.Add($scriptPathButton)

$scriptPathStatus = New-Object System.Windows.Forms.Label
$scriptPathStatus.Location = New-Object System.Drawing.Point(715, ($yPos + 2))
$scriptPathStatus.Size = New-Object System.Drawing.Size(20, 20)
$scriptPathStatus.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
$mainTab.Controls.Add($scriptPathStatus)

$yPos += 40

# Input CSV Section
$inputCSVLabel = New-Object System.Windows.Forms.Label
$inputCSVLabel.Text = "Input CSV:"
$inputCSVLabel.Location = New-Object System.Drawing.Point(10, $yPos)
$inputCSVLabel.Size = New-Object System.Drawing.Size(100, 20)
$mainTab.Controls.Add($inputCSVLabel)

$inputCSVTextBox = New-Object System.Windows.Forms.TextBox
$inputCSVTextBox.Text = $config.CSVPath
$inputCSVTextBox.Location = New-Object System.Drawing.Point(120, $yPos)
$inputCSVTextBox.Size = New-Object System.Drawing.Size(500, 20)
$mainTab.Controls.Add($inputCSVTextBox)

$inputCSVButton = New-Object System.Windows.Forms.Button
$inputCSVButton.Text = "Browse"
$inputCSVButton.Location = New-Object System.Drawing.Point(630, $yPos)
$inputCSVButton.Size = New-Object System.Drawing.Size(75, 23)
$inputCSVButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        if (Test-Path $inputCSVTextBox.Text) {
            $openFileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($inputCSVTextBox.Text)
        }
        $openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $inputCSVTextBox.Text = $openFileDialog.FileName
            Update-ValidationStatus $inputCSVTextBox $inputCSVStatus "CSV"
        }
    })
$mainTab.Controls.Add($inputCSVButton)

$inputCSVStatus = New-Object System.Windows.Forms.Label
$inputCSVStatus.Location = New-Object System.Drawing.Point(715, ($yPos + 2))
$inputCSVStatus.Size = New-Object System.Drawing.Size(20, 20)
$inputCSVStatus.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
$mainTab.Controls.Add($inputCSVStatus)

$yPos += 30

# CSV Preview Button
$previewCSVButton = New-Object System.Windows.Forms.Button
$previewCSVButton.Text = "Preview CSV Data"
$previewCSVButton.Location = New-Object System.Drawing.Point(120, $yPos)
$previewCSVButton.Size = New-Object System.Drawing.Size(120, 25)
$previewCSVButton.Add_Click({
        if (Test-Path $inputCSVTextBox.Text) {
            try {
                $csvData = Import-Csv -Path $inputCSVTextBox.Text | Select-Object -First 10
                $preview = $csvData | ConvertTo-Csv -NoTypeInformation | Out-String
                [System.Windows.Forms.MessageBox]::Show("First 10 rows preview:`n`n$preview", "CSV Preview", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to preview CSV: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Please select a valid CSV file first.", "No File Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })
$mainTab.Controls.Add($previewCSVButton)

$yPos += 40

# Output Folder Section
$outputFolderLabel = New-Object System.Windows.Forms.Label
$outputFolderLabel.Text = "Output Folder:"
$outputFolderLabel.Location = New-Object System.Drawing.Point(10, $yPos)
$outputFolderLabel.Size = New-Object System.Drawing.Size(100, 20)
$mainTab.Controls.Add($outputFolderLabel)

$outputFolderTextBox = New-Object System.Windows.Forms.TextBox
$outputFolderTextBox.Text = $config.OutputFolderPath
$outputFolderTextBox.Location = New-Object System.Drawing.Point(120, $yPos)
$outputFolderTextBox.Size = New-Object System.Drawing.Size(500, 20)
$mainTab.Controls.Add($outputFolderTextBox)

$outputFolderButton = New-Object System.Windows.Forms.Button
$outputFolderButton.Text = "Browse"
$outputFolderButton.Location = New-Object System.Drawing.Point(630, $yPos)
$outputFolderButton.Size = New-Object System.Drawing.Size(75, 23)
$outputFolderButton.Add_Click({
        $folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowserDialog.SelectedPath = $outputFolderTextBox.Text
        if ($folderBrowserDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $outputFolderTextBox.Text = $folderBrowserDialog.SelectedPath
            Update-ValidationStatus $outputFolderTextBox $outputFolderStatus "Folder"
        }
    })
$mainTab.Controls.Add($outputFolderButton)

$outputFolderStatus = New-Object System.Windows.Forms.Label
$outputFolderStatus.Location = New-Object System.Drawing.Point(715, ($yPos + 2))
$outputFolderStatus.Size = New-Object System.Drawing.Size(20, 20)
$outputFolderStatus.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
$mainTab.Controls.Add($outputFolderStatus)

$yPos += 40

# Test Mode Section
$testCheckBox = New-Object System.Windows.Forms.CheckBox
$testCheckBox.Text = "Test Mode (WhatIf - no actual changes will be made)"
$testCheckBox.Location = New-Object System.Drawing.Point(10, $yPos)
$testCheckBox.Size = New-Object System.Drawing.Size(400, 25)
$testCheckBox.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$testCheckBox.ForeColor = [System.Drawing.Color]::DarkBlue
$mainTab.Controls.Add($testCheckBox)

$yPos += 40

# Credentials Section
$credentialsGroupBox = New-Object System.Windows.Forms.GroupBox
$credentialsGroupBox.Text = "Domain Credentials"
$credentialsGroupBox.Location = New-Object System.Drawing.Point(10, $yPos)
$credentialsGroupBox.Size = New-Object System.Drawing.Size(720, 80)
$mainTab.Controls.Add($credentialsGroupBox)

$credentialButton = New-Object System.Windows.Forms.Button
$credentialButton.Text = "Enter Foreign Domain Credentials"
$credentialButton.Location = New-Object System.Drawing.Point(20, 25)
$credentialButton.Size = New-Object System.Drawing.Size(250, 25)
$credentialButton.Add_Click({
        try {
            $script:foreignAdminCreds = Get-Credential -Message "Enter Foreign Domain Admin Credentials"
            if ($script:foreignAdminCreds) {
                $credentialStatusLabel.Text = "OK - Credentials entered for: $($script:foreignAdminCreds.UserName)"
                $credentialStatusLabel.ForeColor = [System.Drawing.Color]::Green
                Write-Log "Credentials entered for user: $($script:foreignAdminCreds.UserName)" "Success"
            }
        }
        catch {
            Write-Log "Failed to get credentials: $($_.Exception.Message)" "Error"
        }
    })
$credentialsGroupBox.Controls.Add($credentialButton)

$credentialStatusLabel = New-Object System.Windows.Forms.Label
$credentialStatusLabel.Text = "No credentials entered"
$credentialStatusLabel.Location = New-Object System.Drawing.Point(20, 55)
$credentialStatusLabel.Size = New-Object System.Drawing.Size(600, 20)
$credentialStatusLabel.ForeColor = [System.Drawing.Color]::Orange
$credentialsGroupBox.Controls.Add($credentialStatusLabel)

$yPos += 100

# Action Buttons Section
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Location = New-Object System.Drawing.Point(10, $yPos)
$buttonPanel.Size = New-Object System.Drawing.Size(720, 80)
$mainTab.Controls.Add($buttonPanel)

$runButton = New-Object System.Windows.Forms.Button
$runButton.Text = "Run Script"
$runButton.Location = New-Object System.Drawing.Point(0, 10)
$runButton.Size = New-Object System.Drawing.Size(100, 35)
$runButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$runButton.BackColor = [System.Drawing.Color]::LightGreen
$buttonPanel.Controls.Add($runButton)

$saveConfigButton = New-Object System.Windows.Forms.Button
$saveConfigButton.Text = "Save Config"
$saveConfigButton.Location = New-Object System.Drawing.Point(110, 10)
$saveConfigButton.Size = New-Object System.Drawing.Size(100, 35)
$saveConfigButton.Add_Click({
        $currentConfig = @{
            CSVPath            = $inputCSVTextBox.Text
            OutputFolderPath   = $outputFolderTextBox.Text
            ScriptPath         = $scriptPathTextBox.Text
            EmailFrom          = $emailFromTextBox.Text
            EmailTo            = $emailToTextBox.Text
            EmailSubject       = $emailSubjectTextBox.Text
            EmailBody          = $emailBodyTextBox.Text
            SMTPServer         = $smtpServerTextBox.Text
            SMTPPort           = $smtpPortTextBox.Text
            MaxRetries         = [int]$maxRetriesTextBox.Text
            RetryDelaySeconds  = [int]$retryDelayTextBox.Text
        }
        Save-Configuration -Config $currentConfig
        Write-Log "Configuration saved successfully" "Success"
    })
$buttonPanel.Controls.Add($saveConfigButton)

$openOutputButton = New-Object System.Windows.Forms.Button
$openOutputButton.Text = "Open Output"
$openOutputButton.Location = New-Object System.Drawing.Point(220, 10)
$openOutputButton.Size = New-Object System.Drawing.Size(100, 35)
$openOutputButton.Enabled = $false
$openOutputButton.Add_Click({
        if (Test-Path $outputFolderTextBox.Text) {
            Start-Process $outputFolderTextBox.Text
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Output folder does not exist", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
$buttonPanel.Controls.Add($openOutputButton)
# === ADVANCED TAB CONTROLS ===
$advYPos = 20

# Advanced Options Group
$advancedGroupBox = New-Object System.Windows.Forms.GroupBox
$advancedGroupBox.Text = "Advanced Script Options"
$advancedGroupBox.Location = New-Object System.Drawing.Point(10, $advYPos)
$advancedGroupBox.Size = New-Object System.Drawing.Size(720, 150)
$advancedTab.Controls.Add($advancedGroupBox)

# Max Retries
$maxRetriesLabel = New-Object System.Windows.Forms.Label
$maxRetriesLabel.Text = "Max Retries:"
$maxRetriesLabel.Location = New-Object System.Drawing.Point(20, 30)
$maxRetriesLabel.Size = New-Object System.Drawing.Size(100, 20)
$advancedGroupBox.Controls.Add($maxRetriesLabel)

$maxRetriesTextBox = New-Object System.Windows.Forms.TextBox
$maxRetriesTextBox.Text = $config.MaxRetries.ToString()
$maxRetriesTextBox.Location = New-Object System.Drawing.Point(130, 30)
$maxRetriesTextBox.Size = New-Object System.Drawing.Size(60, 20)
$advancedGroupBox.Controls.Add($maxRetriesTextBox)

# Retry Delay
$retryDelayLabel = New-Object System.Windows.Forms.Label
$retryDelayLabel.Text = "Retry Delay (sec):"
$retryDelayLabel.Location = New-Object System.Drawing.Point(220, 30)
$retryDelayLabel.Size = New-Object System.Drawing.Size(120, 20)
$advancedGroupBox.Controls.Add($retryDelayLabel)

$retryDelayTextBox = New-Object System.Windows.Forms.TextBox
$retryDelayTextBox.Text = $config.RetryDelaySeconds.ToString()
$retryDelayTextBox.Location = New-Object System.Drawing.Point(350, 30)
$retryDelayTextBox.Size = New-Object System.Drawing.Size(60, 20)
$advancedGroupBox.Controls.Add($retryDelayTextBox)

# Verbose Logging
$verboseCheckBox = New-Object System.Windows.Forms.CheckBox
$verboseCheckBox.Text = "Enable Verbose Logging"
$verboseCheckBox.Location = New-Object System.Drawing.Point(20, 70)
$verboseCheckBox.Size = New-Object System.Drawing.Size(200, 25)
$verboseCheckBox.Checked = $true
$advancedGroupBox.Controls.Add($verboseCheckBox)

$advYPos += 170

# Performance Monitoring Group
$perfGroupBox = New-Object System.Windows.Forms.GroupBox
$perfGroupBox.Text = "Performance Monitoring"
$perfGroupBox.Location = New-Object System.Drawing.Point(10, $advYPos)
$perfGroupBox.Size = New-Object System.Drawing.Size(720, 100)
$advancedTab.Controls.Add($perfGroupBox)

$perfStatsLabel = New-Object System.Windows.Forms.Label
$perfStatsLabel.Text = "Performance statistics will be shown here after script execution"
$perfStatsLabel.Location = New-Object System.Drawing.Point(20, 30)
$perfStatsLabel.Size = New-Object System.Drawing.Size(680, 60)
$perfStatsLabel.ForeColor = [System.Drawing.Color]::Gray
$perfGroupBox.Controls.Add($perfStatsLabel)

# === EMAIL TAB CONTROLS ===
$emailYPos = 20

# Email Configuration Section
$emailConfigGroupBox = New-Object System.Windows.Forms.GroupBox
$emailConfigGroupBox.Text = "Email Notification Settings"
$emailConfigGroupBox.Location = New-Object System.Drawing.Point(10, $emailYPos)
$emailConfigGroupBox.Size = New-Object System.Drawing.Size(720, 300)
$emailTab.Controls.Add($emailConfigGroupBox)

# Email From
$emailFromLabel = New-Object System.Windows.Forms.Label
$emailFromLabel.Text = "From:"
$emailFromLabel.Location = New-Object System.Drawing.Point(20, 30)
$emailFromLabel.Size = New-Object System.Drawing.Size(100, 20)
$emailConfigGroupBox.Controls.Add($emailFromLabel)

$emailFromTextBox = New-Object System.Windows.Forms.TextBox
$emailFromTextBox.Text = $config.EmailFrom
$emailFromTextBox.Location = New-Object System.Drawing.Point(130, 30)
$emailFromTextBox.Size = New-Object System.Drawing.Size(300, 20)
$emailConfigGroupBox.Controls.Add($emailFromTextBox)

$emailFromStatus = New-Object System.Windows.Forms.Label
$emailFromStatus.Location = New-Object System.Drawing.Point(440, 32)
$emailFromStatus.Size = New-Object System.Drawing.Size(20, 20)
$emailFromStatus.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
$emailConfigGroupBox.Controls.Add($emailFromStatus)

# Email To
$emailToLabel = New-Object System.Windows.Forms.Label
$emailToLabel.Text = "To:"
$emailToLabel.Location = New-Object System.Drawing.Point(20, 60)
$emailToLabel.Size = New-Object System.Drawing.Size(100, 20)
$emailConfigGroupBox.Controls.Add($emailToLabel)

$emailToTextBox = New-Object System.Windows.Forms.TextBox
$emailToTextBox.Text = $config.EmailTo
$emailToTextBox.Location = New-Object System.Drawing.Point(130, 60)
$emailToTextBox.Size = New-Object System.Drawing.Size(300, 20)
$emailConfigGroupBox.Controls.Add($emailToTextBox)

$emailToStatus = New-Object System.Windows.Forms.Label
$emailToStatus.Location = New-Object System.Drawing.Point(440, 62)
$emailToStatus.Size = New-Object System.Drawing.Size(20, 20)
$emailToStatus.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
$emailConfigGroupBox.Controls.Add($emailToStatus)

# Email Subject
$emailSubjectLabel = New-Object System.Windows.Forms.Label
$emailSubjectLabel.Text = "Subject:"
$emailSubjectLabel.Location = New-Object System.Drawing.Point(20, 90)
$emailSubjectLabel.Size = New-Object System.Drawing.Size(100, 20)
$emailConfigGroupBox.Controls.Add($emailSubjectLabel)

$emailSubjectTextBox = New-Object System.Windows.Forms.TextBox
$emailSubjectTextBox.Text = $config.EmailSubject
$emailSubjectTextBox.Location = New-Object System.Drawing.Point(130, 90)
$emailSubjectTextBox.Size = New-Object System.Drawing.Size(500, 20)
$emailConfigGroupBox.Controls.Add($emailSubjectTextBox)

# Email Body
$emailBodyLabel = New-Object System.Windows.Forms.Label
$emailBodyLabel.Text = "Body:"
$emailBodyLabel.Location = New-Object System.Drawing.Point(20, 120)
$emailBodyLabel.Size = New-Object System.Drawing.Size(100, 20)
$emailConfigGroupBox.Controls.Add($emailBodyLabel)

$emailBodyTextBox = New-Object System.Windows.Forms.TextBox
$emailBodyTextBox.Text = $config.EmailBody
$emailBodyTextBox.Location = New-Object System.Drawing.Point(130, 120)
$emailBodyTextBox.Size = New-Object System.Drawing.Size(500, 80)
$emailBodyTextBox.Multiline = $true
$emailBodyTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$emailConfigGroupBox.Controls.Add($emailBodyTextBox)

# SMTP Settings
$smtpLabel = New-Object System.Windows.Forms.Label
$smtpLabel.Text = "SMTP Server:"
$smtpLabel.Location = New-Object System.Drawing.Point(20, 210)
$smtpLabel.Size = New-Object System.Drawing.Size(100, 20)
$emailConfigGroupBox.Controls.Add($smtpLabel)

$smtpServerTextBox = New-Object System.Windows.Forms.TextBox
$smtpServerTextBox.Text = $config.SMTPServer
$smtpServerTextBox.Location = New-Object System.Drawing.Point(130, 210)
$smtpServerTextBox.Size = New-Object System.Drawing.Size(200, 20)
$emailConfigGroupBox.Controls.Add($smtpServerTextBox)

$smtpPortLabel = New-Object System.Windows.Forms.Label
$smtpPortLabel.Text = "Port:"
$smtpPortLabel.Location = New-Object System.Drawing.Point(350, 210)
$smtpPortLabel.Size = New-Object System.Drawing.Size(50, 20)
$emailConfigGroupBox.Controls.Add($smtpPortLabel)

$smtpPortTextBox = New-Object System.Windows.Forms.TextBox
$smtpPortTextBox.Text = $config.SMTPPort
$smtpPortTextBox.Location = New-Object System.Drawing.Point(410, 210)
$smtpPortTextBox.Size = New-Object System.Drawing.Size(80, 20)
$emailConfigGroupBox.Controls.Add($smtpPortTextBox)

# Email Test Button
$testEmailButton = New-Object System.Windows.Forms.Button
$testEmailButton.Text = "Test Email Settings"
$testEmailButton.Location = New-Object System.Drawing.Point(20, 250)
$testEmailButton.Size = New-Object System.Drawing.Size(150, 25)
$testEmailButton.Add_Click({
        try {
            $emailParams = @{
                From       = $emailFromTextBox.Text
                To         = $emailToTextBox.Text
                Subject    = "Test Email - $($emailSubjectTextBox.Text)"
                Body       = "This is a test email from the AD Group Management Tool.`n`nTimestamp: $(Get-Date)"
                SmtpServer = $smtpServerTextBox.Text
            }

            if ($smtpPortTextBox.Text) {
                $emailParams.Port = [int]$smtpPortTextBox.Text
            }

            Send-MailMessage @emailParams
            [System.Windows.Forms.MessageBox]::Show("Test email sent successfully!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            Write-Log "Test email sent successfully" "Success"
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to send test email: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            Write-Log "Failed to send test email: $($_.Exception.Message)" "Error"
        }
    })
$emailConfigGroupBox.Controls.Add($testEmailButton)

# === LOGS TAB CONTROLS ===
$logsYPos = 20

# Real-time Logs
$logsLabel = New-Object System.Windows.Forms.Label
$logsLabel.Text = "Real-time Execution Logs:"
$logsLabel.Location = New-Object System.Drawing.Point(10, $logsYPos)
$logsLabel.Size = New-Object System.Drawing.Size(200, 20)
$logsTab.Controls.Add($logsLabel)

$script:logTextBox = New-Object System.Windows.Forms.TextBox
$script:logTextBox.Location = New-Object System.Drawing.Point(10, ($logsYPos + 25))
$script:logTextBox.Size = New-Object System.Drawing.Size(720, 300)
$script:logTextBox.Multiline = $true
$script:logTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$script:logTextBox.ReadOnly = $true
$script:logTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$script:logTextBox.BackColor = [System.Drawing.Color]::Black
$script:logTextBox.ForeColor = [System.Drawing.Color]::White
$logsTab.Controls.Add($script:logTextBox)

$logsYPos += 340

# Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, $logsYPos)
$progressBar.Size = New-Object System.Drawing.Size(720, 25)
$progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$logsTab.Controls.Add($progressBar)

$logsYPos += 40

# Results Summary
$resultsGroupBox = New-Object System.Windows.Forms.GroupBox
$resultsGroupBox.Text = "Execution Results Summary"
$resultsGroupBox.Location = New-Object System.Drawing.Point(10, $logsYPos)
$resultsGroupBox.Size = New-Object System.Drawing.Size(720, 120)
$logsTab.Controls.Add($resultsGroupBox)

$resultsLabel = New-Object System.Windows.Forms.Label
$resultsLabel.Text = "No execution results yet"
$resultsLabel.Location = New-Object System.Drawing.Point(20, 30)
$resultsLabel.Size = New-Object System.Drawing.Size(680, 80)
$resultsLabel.ForeColor = [System.Drawing.Color]::Gray
$resultsGroupBox.Controls.Add($resultsLabel)

# Control Buttons
$logsYPos += 140

$clearLogsButton = New-Object System.Windows.Forms.Button
$clearLogsButton.Text = "Clear Logs"
$clearLogsButton.Location = New-Object System.Drawing.Point(10, $logsYPos)
$clearLogsButton.Size = New-Object System.Drawing.Size(100, 25)
$clearLogsButton.Add_Click({
        $script:logTextBox.Clear()
        Write-Log "Logs cleared" "Info"
    })
$logsTab.Controls.Add($clearLogsButton)

$exportLogsButton = New-Object System.Windows.Forms.Button
$exportLogsButton.Text = "Export Logs"
$exportLogsButton.Location = New-Object System.Drawing.Point(120, $logsYPos)
$exportLogsButton.Size = New-Object System.Drawing.Size(100, 25)
$exportLogsButton.Add_Click({
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        $saveFileDialog.FileName = "ADGroupScript_Logs_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:logTextBox.Text | Out-File -FilePath $saveFileDialog.FileName -Encoding UTF8
            Write-Log "Logs exported to: $($saveFileDialog.FileName)" "Success"
        }
    })
$logsTab.Controls.Add($exportLogsButton)

# Enhanced Run Button Click Event
$runButton.Add_Click({
        if ($script:isRunning) {
            [System.Windows.Forms.MessageBox]::Show("Script is already running. Please wait for it to complete.", "Already Running", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $script:isRunning = $true
        $runButton.Enabled = $false
        $statusLabel.Text = "Status: Validating inputs..."
        $statusLabel.ForeColor = [System.Drawing.Color]::Orange

        try {
            # Comprehensive validation
            $validationErrors = @()

            # Validate script path
            if ([string]::IsNullOrWhiteSpace($scriptPathTextBox.Text) -or -not (Test-Path $scriptPathTextBox.Text)) {
                $validationErrors += "Invalid script path. Please select the Add-ForeignUsertoFileShareADGroup.ps1 file."
            }

            # Validate CSV
            $csvIssues = Validate-CSVFile -CsvPath $inputCSVTextBox.Text
            if ($csvIssues.Count -gt 0) {
                $criticalIssues = $csvIssues | Where-Object { -not ($_ -like "Warning*") }
                if ($criticalIssues.Count -gt 0) {
                    $validationErrors += $criticalIssues
                }
            }

            # Validate output folder
            if ([string]::IsNullOrWhiteSpace($outputFolderTextBox.Text)) {
                $validationErrors += "Output folder is required."
            }
            elseif (-not (Test-Path $outputFolderTextBox.Text)) {
                try {
                    New-Item -Path $outputFolderTextBox.Text -ItemType Directory -Force | Out-Null
                    Write-Log "Created output directory: $($outputFolderTextBox.Text)" "Info"
                }
                catch {
                    $validationErrors += "Cannot create output folder: $($_.Exception.Message)"
                }
            }

            # Validate credentials
            if (-not $script:foreignAdminCreds) {
                $response = [System.Windows.Forms.MessageBox]::Show("No foreign domain credentials entered. Do you want to continue without them?", "Missing Credentials", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
                if ($response -eq [System.Windows.Forms.DialogResult]::No) {
                    $validationErrors += "Foreign domain credentials are required."
                }
            }

            # Validate advanced options
            if (-not ($maxRetriesTextBox.Text -match "^\d+$") -or [int]$maxRetriesTextBox.Text -lt 1 -or [int]$maxRetriesTextBox.Text -gt 10) {
                $validationErrors += "Max Retries must be a number between 1 and 10."
            }

            if (-not ($retryDelayTextBox.Text -match "^\d+$") -or [int]$retryDelayTextBox.Text -lt 1 -or [int]$retryDelayTextBox.Text -gt 60) {
                $validationErrors += "Retry Delay must be a number between 1 and 60 seconds."
            }

            if ($validationErrors.Count -gt 0) {
                $errorMessage = "Validation failed:`n`n" + ($validationErrors -join "`n")
                [System.Windows.Forms.MessageBox]::Show($errorMessage, "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                Write-Log "Validation failed: $($validationErrors -join '; ')" "Error"
                return
            }

            Write-Log "Validation passed. Starting script execution..." "Success"
            $statusLabel.Text = "Status: Running script..."
            $statusLabel.ForeColor = [System.Drawing.Color]::Blue

            # Switch to logs tab
            $tabControl.SelectedTab = $logsTab

            # Prepare script parameters
            $scriptParams = @{
                InputCsvPath      = $inputCSVTextBox.Text
                OutputFolderPath  = $outputFolderTextBox.Text
                MaxRetries        = [int]$maxRetriesTextBox.Text
                RetryDelaySeconds = [int]$retryDelayTextBox.Text
            }

            if ($testCheckBox.Checked) {
                $scriptParams.Test = $true
                Write-Log "Running in TEST MODE - no actual changes will be made" "Warning"
            }

            if ($script:foreignAdminCreds) {
                $scriptParams.ForeignAdminCreds = $script:foreignAdminCreds
                Write-Log "Using foreign domain credentials for: $($script:foreignAdminCreds.UserName)" "Info"
            }

            # Note: For switch parameters in splatting, we only add them if they should be true
            if ($verboseCheckBox.Checked) {
                $scriptParams.Verbose = $true
                Write-Log "Verbose output enabled" "Info"
            }

            Write-Log "Executing script: $($scriptPathTextBox.Text)" "Info"
            Write-Log "Parameters: $($scriptParams | ConvertTo-Json -Compress)" "Info"

            # Execute the script
            $startTime = Get-Date
            $scriptSuccess = $false

            try {
                # Change to script directory
                $scriptDir = [System.IO.Path]::GetDirectoryName($scriptPathTextBox.Text)
                Push-Location $scriptDir

                # Execute with progress monitoring
                $progressBar.Value = 10
                Write-Log "Script execution started..." "Info"

                # Execute the script directly to preserve credential context
                Write-Log "Script execution started..." "Info"
                $progressBar.Value = 30

                try {
                    # Change to script directory and execute directly
                    $scriptDir = [System.IO.Path]::GetDirectoryName($scriptPathTextBox.Text)
                    Push-Location $scriptDir

                    # Execute the script with splatting
                    & $scriptPathTextBox.Text @scriptParams

                    Pop-Location
                    $scriptSuccess = $true
                }
                catch {
                    Pop-Location
                    throw $_.Exception
                }

                $endTime = Get-Date
                $duration = ($endTime - $startTime).TotalSeconds

                $progressBar.Value = 90

                if ($scriptSuccess) {
                    Write-Log "Script completed successfully in $([math]::Round($duration, 2)) seconds" "Success"
                    $statusLabel.Text = "Status: Completed Successfully"
                    $statusLabel.ForeColor = [System.Drawing.Color]::Green

                    # Update results summary
                    $resultsLabel.Text = "Execution completed successfully!`nDuration: $([math]::Round($duration, 2)) seconds`nMode: $(if ($testCheckBox.Checked) { 'TEST' } else { 'LIVE' })`nCheck output folder for detailed logs."
                    $resultsLabel.ForeColor = [System.Drawing.Color]::Green

                    # Enable output folder button
                    $openOutputButton.Enabled = $true

                    # Send email if configured
                    if ($emailFromTextBox.Text -and $emailToTextBox.Text -and $smtpServerTextBox.Text) {
                        try {
                            Write-Log "Sending completion email..." "Info"

                            # Look for the specific ADUserToGroupLog files with correct naming pattern
                            $attachmentPath = Join-Path $outputFolderTextBox.Text "ADUserToGroupLog_*.csv"
                            $attachments = Get-ChildItem $attachmentPath -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1

                            $emailParams = @{
                                From       = $emailFromTextBox.Text
                                To         = $emailToTextBox.Text
                                Subject    = $emailSubjectTextBox.Text
                                Body       = "$($emailBodyTextBox.Text)`n`nExecution Summary:`nDuration: $([math]::Round($duration, 2)) seconds`nMode: $(if ($testCheckBox.Checked) { 'TEST' } else { 'LIVE' })`nCompleted: $(Get-Date)"
                                SmtpServer = $smtpServerTextBox.Text
                            }

                            if ($smtpPortTextBox.Text) {
                                $emailParams.Port = [int]$smtpPortTextBox.Text
                            }

                            if ($attachments) {
                                $emailParams.Attachments = $attachments.FullName
                                Write-Log "Attaching log file: $($attachments.Name)" "Info"
                            }

                            Send-MailMessage @emailParams
                            Write-Log "Email notification sent successfully" "Success"
                            $statusLabel.Text = "Status: Completed - Email Sent"
                        }
                        catch {
                            Write-Log "Failed to send email notification: $($_.Exception.Message)" "Error"
                            $statusLabel.Text = "Status: Completed - Email Failed"
                        }
                    }
                }
                else {
                    Write-Log "Script execution failed" "Error"
                    $statusLabel.Text = "Status: Failed"
                    $statusLabel.ForeColor = [System.Drawing.Color]::Red
                    $resultsLabel.Text = "Execution failed. Check logs for details."
                    $resultsLabel.ForeColor = [System.Drawing.Color]::Red
                }

                $progressBar.Value = 100
            }
            catch {
                Write-Log "Script execution error: $($_.Exception.Message)" "Error"
                $statusLabel.Text = "Status: Error"
                $statusLabel.ForeColor = [System.Drawing.Color]::Red
                $resultsLabel.Text = "Execution error: $($_.Exception.Message)"
                $resultsLabel.ForeColor = [System.Drawing.Color]::Red

                [System.Windows.Forms.MessageBox]::Show("Script execution failed: $($_.Exception.Message)", "Execution Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
            finally {
                Pop-Location -ErrorAction SilentlyContinue
            }
        }
        catch {
            Write-Log "Unexpected error: $($_.Exception.Message)" "Error"
            $statusLabel.Text = "Status: Error"
            $statusLabel.ForeColor = [System.Drawing.Color]::Red
            [System.Windows.Forms.MessageBox]::Show("An unexpected error occurred: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
        finally {
            $script:isRunning = $false
            $runButton.Enabled = $true
            $progressBar.Value = 0
        }
    })

# Add validation event handlers
$inputCSVTextBox.Add_TextChanged({ Update-ValidationStatus $inputCSVTextBox $inputCSVStatus "CSV" })
$outputFolderTextBox.Add_TextChanged({ Update-ValidationStatus $outputFolderTextBox $outputFolderStatus "Folder" })
$scriptPathTextBox.Add_TextChanged({ Update-ValidationStatus $scriptPathTextBox $scriptPathStatus "File" })
$emailFromTextBox.Add_TextChanged({ Update-ValidationStatus $emailFromTextBox $emailFromStatus "Email" })
$emailToTextBox.Add_TextChanged({ Update-ValidationStatus $emailToTextBox $emailToStatus "Email" })
$maxRetriesTextBox.Add_TextChanged({ Update-ValidationStatus $maxRetriesTextBox $null "Number" })
$retryDelayTextBox.Add_TextChanged({ Update-ValidationStatus $retryDelayTextBox $null "Number" })

# Form Load Event - populate with saved configuration
$form.Add_Shown({
        Write-Log "Enhanced AD Group Management Tool v2.0 started" "Info"
        Write-Log "Loading configuration..." "Info"

        # Populate fields with configuration
        $inputCSVTextBox.Text = $config.CSVPath
        $outputFolderTextBox.Text = $config.OutputFolderPath
        $scriptPathTextBox.Text = $config.ScriptPath
        $emailFromTextBox.Text = $config.EmailFrom
        $emailToTextBox.Text = $config.EmailTo
        $emailSubjectTextBox.Text = $config.EmailSubject
        $emailBodyTextBox.Text = $config.EmailBody
        $smtpServerTextBox.Text = $config.SMTPServer
        $smtpPortTextBox.Text = $config.SMTPPort
        $maxRetriesTextBox.Text = $config.MaxRetries.ToString()
        $retryDelayTextBox.Text = $config.RetryDelaySeconds.ToString()

        # Trigger validation
        Start-Sleep -Milliseconds 100
        Update-ValidationStatus $inputCSVTextBox $inputCSVStatus "CSV"
        Update-ValidationStatus $outputFolderTextBox $outputFolderStatus "Folder"
        Update-ValidationStatus $scriptPathTextBox $scriptPathStatus "File"
        Update-ValidationStatus $emailFromTextBox $emailFromStatus "Email"
        Update-ValidationStatus $emailToTextBox $emailToStatus "Email"

        Write-Log "Configuration loaded and validation completed" "Success"
        Write-Log "Ready to process AD group assignments" "Info"
    })

# Form Closing Event - save configuration
$form.Add_FormClosing({
        if (-not $script:isRunning) {
            $currentConfig = @{
                CSVPath            = $inputCSVTextBox.Text
                OutputFolderPath   = $outputFolderTextBox.Text
                ScriptPath         = $scriptPathTextBox.Text
                EmailFrom          = $emailFromTextBox.Text
                EmailTo            = $emailToTextBox.Text
                EmailSubject       = $emailSubjectTextBox.Text
                EmailBody          = $emailBodyTextBox.Text
                SMTPServer         = $smtpServerTextBox.Text
                SMTPPort           = $smtpPortTextBox.Text
                MaxRetries         = if ($maxRetriesTextBox.Text -match "^\d+$") { [int]$maxRetriesTextBox.Text } else { 3 }
                RetryDelaySeconds  = if ($retryDelayTextBox.Text -match "^\d+$") { [int]$retryDelayTextBox.Text } else { 5 }
            }
            Save-Configuration -Config $currentConfig
            Write-Log "Configuration saved on exit" "Info"
        }
        else {
            $_.Cancel = $true
            [System.Windows.Forms.MessageBox]::Show("Cannot close while script is running. Please wait for completion.", "Script Running", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })

# Show the enhanced form
Write-Host "Starting Enhanced AD Group Management Tool..." -ForegroundColor Green
$form.ShowDialog()