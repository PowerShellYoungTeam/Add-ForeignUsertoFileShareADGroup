Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$defaultCSVPath = "C:\temp\input.csv"
$defaultOutputFolderPath = "C:\temp\"

$form = New-Object System.Windows.Forms.Form
$form.Text = "Add-ForeignUsertoFileShareADGroup GUI Controller"
$form.Size = New-Object System.Drawing.Size(600, 400)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Status: Not Started"
$statusLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
$statusLabel.Location = New-Object System.Drawing.Point(10, 10)
$statusLabel.Size = New-Object System.Drawing.Size(500, 30)
$form.Controls.Add($statusLabel)

$inputCSVLabel = New-Object System.Windows.Forms.Label
$inputCSVLabel.Text = "Input CSV Path:"
$inputCSVLabel.Location = New-Object System.Drawing.Point(10, 50)
$form.Controls.Add($inputCSVLabel)

$inputCSVTextBox = New-Object System.Windows.Forms.TextBox
$inputCSVTextBox.Text = "Choose file"
$inputCSVTextBox.Location = New-Object System.Drawing.Point(120, 50)
$inputCSVTextBox.Size = New-Object System.Drawing.Size(350, 20)
$form.Controls.Add($inputCSVTextBox)

$inputCSVButton = New-Object System.Windows.Forms.Button
$inputCSVButton.Text = "Browse"
$inputCSVButton.Location = New-Object System.Drawing.Point(480, 50)
$inputCSVButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($defaultCSVPath)
        $openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $inputCSVTextBox.Text = $openFileDialog.FileName
        }
    })
$form.Controls.Add($inputCSVButton)

$outputFolderLabel = New-Object System.Windows.Forms.Label
$outputFolderLabel.Text = "Output Folder:"
$outputFolderLabel.Location = New-Object System.Drawing.Point(10, 90)
$outputFolderLabel.Size = New-Object System.Drawing.Size(110, 20)  # Increased size
$form.Controls.Add($outputFolderLabel)

$outputFolderTextBox = New-Object System.Windows.Forms.TextBox
$outputFolderTextBox.Text = "Choose folder"
$outputFolderTextBox.Location = New-Object System.Drawing.Point(120, 90)
$outputFolderTextBox.Size = New-Object System.Drawing.Size(350, 20)
$form.Controls.Add($outputFolderTextBox)

$outputFolderButton = New-Object System.Windows.Forms.Button
$outputFolderButton.Text = "Browse"
$outputFolderButton.Location = New-Object System.Drawing.Point(480, 90)
$outputFolderButton.Add_Click({
        $folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowserDialog.SelectedPath = $defaultOutputFolderPath
        if ($folderBrowserDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $outputFolderTextBox.Text = $folderBrowserDialog.SelectedPath
        }
    })
$form.Controls.Add($outputFolderButton)

$testCheckBox = New-Object System.Windows.Forms.CheckBox
$testCheckBox.Text = "Test - check this if you want to run without making changes (-whatif mode)"
$testCheckBox.Location = New-Object System.Drawing.Point(10, 130)
$form.Controls.Add($testCheckBox)

$credentialButton = New-Object System.Windows.Forms.Button
$credentialButton.Text = "Enter Credentials"
$credentialButton.Location = New-Object System.Drawing.Point(10, 170)
$credentialButton.Add_Click({
        $foreignAdminCreds = Get-Credential -Message "Enter Foreign Admin Credentials"
    })
$form.Controls.Add($credentialButton)

$runButton = New-Object System.Windows.Forms.Button
$runButton.Text = "Run"
$runButton.Location = New-Object System.Drawing.Point(10, 210)
$runButton.Add_Click({
        $statusLabel.Text = "Status: Running"
        # Validate CSV
        if (-not (Test-Path $inputCSVTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Invalid CSV Path")
            $statusLabel.Text = "Status: ERROR"
            return
        }
        try {
            # Run the script
            # HACK *** encase working dir is not the same place as where scripts are stored
            # set-location '\\server\folders\'
            .\add-foreignUsertoFileShareADGroup.ps1 -InputCSVPath $inputCSVTextBox.Text -OutputFolderPath $outputFolderTextBox.Text -ForeignAdminCreds $foreignAdminCreds -Test $testCheckBox.Checked
            $statusLabel.Text = "Status: Complete"
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("An error occurred while running the script: $_")
            $statusLabel.Text = "Status: ERROR"
        }
    })
$form.Controls.Add($runButton)

$openOutputFolderButton = New-Object System.Windows.Forms.Button
$openOutputFolderButton.Text = "Open Output Folder"
$openOutputFolderButton.Location = New-Object System.Drawing.Point(100, 210)
$openOutputFolderButton.Enabled = $false
$openOutputFolderButton.Add_Click({
        Start-Process $outputFolderTextBox.Text
    })
$form.Controls.Add($openOutputFolderButton)

$form.Add_Shown({
        $inputCSVTextBox.Text = $defaultCSVPath
        $outputFolderTextBox.Text = $defaultOutputFolderPath
    })

$form.ShowDialog()