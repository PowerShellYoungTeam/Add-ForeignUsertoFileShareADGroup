Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$defaultCSVPath = "C:\temp\Powershell\BarryGroups.csv"
$defaultOutputFolderPath = "C:\temp\PowerShell"

# Default email settings
$defaultEmailFrom = "noreply@company.com"
$defaultEmailTo = "admin@company.com"
$defaultEmailSubject = "AD Group Script Completion"
$defaultEmailBody = "The Add-ForeignUsertoFileShareADGroup script has completed successfully. Please see the attached output file for details."
$defaultSMTPServer = "smtp.company.com"
$defaultSMTPPort = "587"

$form = New-Object System.Windows.Forms.Form
$form.Text = "Add-ForeignUsertoFileShareADGroup GUI Controller"
$form.Size = New-Object System.Drawing.Size(600, 600)

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

# Email Configuration Section
$emailLabel = New-Object System.Windows.Forms.Label
$emailLabel.Text = "Email Configuration"
$emailLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$emailLabel.Location = New-Object System.Drawing.Point(10, 160)
$emailLabel.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($emailLabel)

$emailFromLabel = New-Object System.Windows.Forms.Label
$emailFromLabel.Text = "From:"
$emailFromLabel.Location = New-Object System.Drawing.Point(10, 190)
$emailFromLabel.Size = New-Object System.Drawing.Size(110, 20)
$form.Controls.Add($emailFromLabel)

$emailFromTextBox = New-Object System.Windows.Forms.TextBox
$emailFromTextBox.Location = New-Object System.Drawing.Point(120, 190)
$emailFromTextBox.Size = New-Object System.Drawing.Size(350, 20)
$form.Controls.Add($emailFromTextBox)

$emailToLabel = New-Object System.Windows.Forms.Label
$emailToLabel.Text = "To:"
$emailToLabel.Location = New-Object System.Drawing.Point(10, 220)
$emailToLabel.Size = New-Object System.Drawing.Size(110, 20)
$form.Controls.Add($emailToLabel)

$emailToTextBox = New-Object System.Windows.Forms.TextBox
$emailToTextBox.Location = New-Object System.Drawing.Point(120, 220)
$emailToTextBox.Size = New-Object System.Drawing.Size(350, 20)
$form.Controls.Add($emailToTextBox)

$emailSubjectLabel = New-Object System.Windows.Forms.Label
$emailSubjectLabel.Text = "Subject:"
$emailSubjectLabel.Location = New-Object System.Drawing.Point(10, 250)
$emailSubjectLabel.Size = New-Object System.Drawing.Size(110, 20)
$form.Controls.Add($emailSubjectLabel)

$emailSubjectTextBox = New-Object System.Windows.Forms.TextBox
$emailSubjectTextBox.Location = New-Object System.Drawing.Point(120, 250)
$emailSubjectTextBox.Size = New-Object System.Drawing.Size(350, 20)
$form.Controls.Add($emailSubjectTextBox)

$emailBodyLabel = New-Object System.Windows.Forms.Label
$emailBodyLabel.Text = "Body:"
$emailBodyLabel.Location = New-Object System.Drawing.Point(10, 280)
$emailBodyLabel.Size = New-Object System.Drawing.Size(110, 20)
$form.Controls.Add($emailBodyLabel)

$emailBodyTextBox = New-Object System.Windows.Forms.TextBox
$emailBodyTextBox.Location = New-Object System.Drawing.Point(120, 280)
$emailBodyTextBox.Size = New-Object System.Drawing.Size(350, 60)
$emailBodyTextBox.Multiline = $true
$emailBodyTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$form.Controls.Add($emailBodyTextBox)

$smtpServerLabel = New-Object System.Windows.Forms.Label
$smtpServerLabel.Text = "SMTP Server:"
$smtpServerLabel.Location = New-Object System.Drawing.Point(10, 350)
$smtpServerLabel.Size = New-Object System.Drawing.Size(110, 20)
$form.Controls.Add($smtpServerLabel)

$smtpServerTextBox = New-Object System.Windows.Forms.TextBox
$smtpServerTextBox.Location = New-Object System.Drawing.Point(120, 350)
$smtpServerTextBox.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($smtpServerTextBox)

$smtpPortLabel = New-Object System.Windows.Forms.Label
$smtpPortLabel.Text = "Port:"
$smtpPortLabel.Location = New-Object System.Drawing.Point(330, 350)
$smtpPortLabel.Size = New-Object System.Drawing.Size(40, 20)
$form.Controls.Add($smtpPortLabel)

$smtpPortTextBox = New-Object System.Windows.Forms.TextBox
$smtpPortTextBox.Location = New-Object System.Drawing.Point(370, 350)
$smtpPortTextBox.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($smtpPortTextBox)

$credentialButton = New-Object System.Windows.Forms.Button
$credentialButton.Text = "Enter Credentials"
$credentialButton.Location = New-Object System.Drawing.Point(10, 380)
$credentialButton.Add_Click({
        $foreignAdminCreds = Get-Credential -Message "Enter Foreign Admin Credentials"
    })
#$form.Controls.Add($credentialButton)

$runButton = New-Object System.Windows.Forms.Button
$runButton.Text = "Run"
$runButton.Location = New-Object System.Drawing.Point(10, 420)
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
            set-location '\\ukgcbpro.uk.gcb.corp\gcbdfs\Data\EUTApplicationSource\_Powershell\AD\'
            .\add-foreignUsertoFileShareADGroup.ps1 -InputCSVPath $inputCSVTextBox.Text -OutputFolderPath $outputFolderTextBox.Text -ForeignAdminCreds $foreignAdminCreds -Test $testCheckBox.Checked

            # Send completion email if email fields are filled
            if ($emailFromTextBox.Text -and $emailToTextBox.Text -and $smtpServerTextBox.Text) {
                try {
                    $attachmentPath = Join-Path $outputFolderTextBox.Text "*.csv"
                    $attachments = Get-ChildItem $attachmentPath -ErrorAction SilentlyContinue

                    $emailParams = @{
                        From       = $emailFromTextBox.Text
                        To         = $emailToTextBox.Text
                        Subject    = $emailSubjectTextBox.Text
                        Body       = $emailBodyTextBox.Text
                        SmtpServer = $smtpServerTextBox.Text
                    }

                    if ($smtpPortTextBox.Text) {
                        $emailParams.Port = [int]$smtpPortTextBox.Text
                    }

                    if ($attachments) {
                        $emailParams.Attachments = $attachments[0].FullName
                    }

                    Send-MailMessage @emailParams
                    $statusLabel.Text = "Status: Complete - Email sent"
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show("Script completed but email failed: $_")
                    $statusLabel.Text = "Status: Complete - Email failed"
                }
            }
            else {
                $statusLabel.Text = "Status: Complete"
            }

            $openOutputFolderButton.Enabled = $true
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("An error occurred while running the script: $_")
            $statusLabel.Text = "Status: ERROR"
        }
    })
$form.Controls.Add($runButton)

$openOutputFolderButton = New-Object System.Windows.Forms.Button
$openOutputFolderButton.Text = "Open Output Folder"
$openOutputFolderButton.Location = New-Object System.Drawing.Point(100, 420)
$openOutputFolderButton.Enabled = $false
$openOutputFolderButton.Add_Click({
        Start-Process $outputFolderTextBox.Text
    })
$form.Controls.Add($openOutputFolderButton)

$form.Add_Shown({
        $inputCSVTextBox.Text = $defaultCSVPath
        $outputFolderTextBox.Text = $defaultOutputFolderPath
        $emailFromTextBox.Text = $defaultEmailFrom
        $emailToTextBox.Text = $defaultEmailTo
        $emailSubjectTextBox.Text = $defaultEmailSubject
        $emailBodyTextBox.Text = $defaultEmailBody
        $smtpServerTextBox.Text = $defaultSMTPServer
        $smtpPortTextBox.Text = $defaultSMTPPort
    })

$form.ShowDialog()