<#
.SYNOPSIS
Enhanced wrapper script to add users from a CSV file to AD groups with comprehensive error handling and validation.

.DESCRIPTION
This script reads user and group information from a CSV file and adds the users to the specified AD groups.
It includes robust validation, retry logic, performance monitoring, and comprehensive logging.
Features include CSV validation and detailed reporting.

.PARAMETER InputCsvPath
The file path to the input CSV file containing user and group information.
Required columns: SourceDomain, SourceUser, TargetDomain, TargetGroup

.PARAMETER OutputFolderPath
The folder path where the log and transcript files will be saved.

.PARAMETER Test
Switch to simulate the operation without making any changes (WhatIf mode).

.PARAMETER ForeignAdminCreds
Credentials for the foreign domain admin. If not provided, will prompt for credentials.

.PARAMETER MaxRetries
Maximum number of retry attempts for failed operations. Default is 3.

.PARAMETER RetryDelaySeconds
Delay in seconds between retry attempts. Default is 5.

.EXAMPLE
.\Add-ForeignUsertoFileShareADGroup.ps1 -InputCsvPath "C:\Input\users.csv" -OutputFolderPath "C:\Output" -Test

.EXAMPLE
.\Add-ForeignUsertoFileShareADGroup.ps1 -InputCsvPath "C:\Input\users.csv" -OutputFolderPath "C:\Output" -MaxRetries 5

.NOTES
Author: Steven Wight with GitHub Copilot
Date: 23/01/2025
Updated: Enhanced for Windows PowerShell 5.1 compatibility
Requires: ActiveDirectory module, appropriate domain permissions

#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$InputCsvPath,

    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$OutputFolderPath,

    [Parameter(Mandatory = $false)]
    [switch]$Test,

    [Parameter(Mandatory = $false)]
    [PSCredential]$ForeignAdminCreds,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 10)]
    [int]$MaxRetries = 3,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 60)]
    [int]$RetryDelaySeconds = 5
)

# FUNCTIONS

function Test-CSVSchema {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$CsvPath
    )

    $requiredColumns = @('SourceDomain', 'SourceUser', 'TargetDomain', 'TargetGroup')

    try {
        $csvData = Import-Csv -Path $CsvPath -ErrorAction Stop
        $csvColumns = $csvData[0].PSObject.Properties.Name

        $missingColumns = $requiredColumns | Where-Object { $_ -notin $csvColumns }

        if ($missingColumns) {
            throw "Missing required columns: $($missingColumns -join ', '). Required columns are: $($requiredColumns -join ', ')"
        }

        # Check for empty rows
        $emptyRows = $csvData | Where-Object {
            [string]::IsNullOrWhiteSpace($_.SourceDomain) -or
            [string]::IsNullOrWhiteSpace($_.SourceUser) -or
            [string]::IsNullOrWhiteSpace($_.TargetDomain) -or
            [string]::IsNullOrWhiteSpace($_.TargetGroup)
        }

        if ($emptyRows) {
            Write-Warning "Found $($emptyRows.Count) rows with missing data. These will be skipped."
        }

        Write-Verbose "CSV schema validation passed. Found $($csvData.Count) total rows."
        return $true
    }
    catch {
        Write-Error "CSV validation failed: $($_.Exception.Message)"
        return $false
    }
}

function Invoke-WithRetry {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,

        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 3,

        [Parameter(Mandatory = $false)]
        [int]$DelaySeconds = 5,

        [Parameter(Mandatory = $false)]
        [string]$OperationName = "Operation"
    )

    $attempt = 1

    do {
        try {
            Write-Verbose "Attempting $OperationName (Attempt $attempt of $($MaxRetries + 1))"
            $result = & $ScriptBlock
            Write-Verbose "$OperationName succeeded on attempt $attempt"
            return $result
        }
        catch {
            Write-Warning "$OperationName failed on attempt $attempt`: $($_.Exception.Message)"

            if ($attempt -le $MaxRetries) {
                Write-Verbose "Waiting $DelaySeconds seconds before retry..."
                Start-Sleep -Seconds $DelaySeconds
            }
            else {
                Write-Error "$OperationName failed after $($MaxRetries + 1) attempts"
                throw
            }
        }
        $attempt++
    } while ($attempt -le ($MaxRetries + 1))
}

function Add-ADUserToGroup {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [PSCustomObject[]]$UserGroupData,

        [Parameter(Mandatory = $true)]
        [string]$LogFilePath,

        [Parameter(Mandatory = $false)]
        [switch]$Test,

        [Parameter(Mandatory = $false)]
        [PSCredential]$ForeignAdminCreds,

        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 3,

        [Parameter(Mandatory = $false)]
        [int]$RetryDelaySeconds = 5
    )

    <#
    .SYNOPSIS
    Enhanced function to add users from source domains to groups in target domains with comprehensive error handling.

    .DESCRIPTION
    This function adds specified users from source domains to specified groups in target domains.
    It includes retry logic, detailed validation, performance monitoring, and comprehensive logging.

    .PARAMETER UserGroupData
    The object array containing user and group information with required properties:
    SourceDomain, SourceUser, TargetDomain, TargetGroup

    .PARAMETER LogFilePath
    The file path where the detailed log will be saved.

    .PARAMETER Test
    Switch to simulate the operation without making changes (WhatIf mode).

    .PARAMETER ForeignAdminCreds
    Credentials for the foreign domain admin.

    .PARAMETER MaxRetries
    Maximum number of retry attempts for failed operations.

    .PARAMETER RetryDelaySeconds
    Delay in seconds between retry attempts.

    .EXAMPLE
    $userGroupData | Add-ADUserToGroup -LogFilePath "C:\Logs\ADUserToGroupLog.csv" -Test -Verbose

    .NOTES
    Author: Steven Wight with GitHub Copilot
    Date: 23/01/2025
    Enhanced with comprehensive error handling and validation
    #>

    begin {
        $logEntries = @()
        $processedCount = 0
        $successCount = 0
        $errorCount = 0
        $startTime = Get-Date
        $allUserGroupData = @()

        # Initialize totalItems - will be calculated from collected data
        $totalItems = 0

        # Only prompt for credentials if not provided
        if (-not $ForeignAdminCreds) {
            Write-Host "Foreign domain credentials required for cross-domain operations..." -ForegroundColor Yellow
            $ForeignAdminCreds = Get-Credential -Message "Enter admin credentials for foreign domain access"
            if (-not $ForeignAdminCreds) {
                throw "Credentials are required for foreign domain access"
            }
        }
        else {
            Write-Verbose "Using provided foreign domain credentials for user: $($ForeignAdminCreds.UserName)"
        }

        # Test AD module availability
        if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
            throw "ActiveDirectory PowerShell module is not available. Please install RSAT tools."
        }

        Import-Module ActiveDirectory -Force -ErrorAction Stop
        Write-Verbose "ActiveDirectory module loaded successfully"
    }

    process {
        # Collect all items first to get accurate count
        foreach ($item in $UserGroupData) {
            $allUserGroupData += $item
        }
    }
    end {
        # Set totalItems from collected data
        $totalItems = $allUserGroupData.Count
        Write-Verbose "Total items from CSV: $totalItems"

        # Process all collected items
        foreach ($entry in $allUserGroupData) {
            $processedCount++
            # Ensure percentage is capped at 100 and converted to integer for Write-Progress
            $percentComplete = if ($totalItems -gt 0) { [math]::Min(100, [math]::Round(($processedCount / $totalItems) * 100)) } else { 0 }

            # Validate entry data
            if ([string]::IsNullOrWhiteSpace($entry.SourceDomain) -or
                [string]::IsNullOrWhiteSpace($entry.SourceUser) -or
                [string]::IsNullOrWhiteSpace($entry.TargetDomain) -or
                [string]::IsNullOrWhiteSpace($entry.TargetGroup)) {
                Write-Warning "Skipping entry with missing data: $($entry | ConvertTo-Json -Compress)"
                continue
            }

            $SourceDomainUser = "$($entry.SourceDomain)\$($entry.SourceUser)"
            $TargetDomainGroup = "$($entry.TargetDomain)\$($entry.TargetGroup)"
            $operationId = [guid]::NewGuid().ToString("N").Substring(0, 8)

            # Calculate detailed percentage for verbose output
            $detailedPercent = if ($totalItems -gt 0) { [math]::Round(($processedCount / $totalItems) * 100, 2) } else { 0 }
            Write-Verbose "[$operationId] Processing: $SourceDomainUser -> $TargetDomainGroup ($processedCount/$totalItems - $detailedPercent%)"
            Write-Progress -Activity "Adding Users to Groups" -Status "Processing $($entry.SourceUser)" -PercentComplete $percentComplete

            if ($PSCmdlet.ShouldProcess("$SourceDomainUser", "Add to group $TargetDomainGroup")) {
                $operationStartTime = Get-Date
                $status = "Error"
                $message = ""

                try {
                    Write-Verbose "[$operationId] Retrieving user $($entry.SourceUser) from source domain $($entry.SourceDomain)"
                    Write-Verbose "[$operationId] Using credentials: $($ForeignAdminCreds.UserName)"

                    if ($Test) {
                        Write-Verbose "[$operationId] TEST MODE: Would add $SourceDomainUser to $TargetDomainGroup"
                        # Get the user object from source domain first
                        $SourceUserObj = Get-ADUser $entry.SourceUser -Server $entry.SourceDomain -Credential $ForeignAdminCreds -ErrorAction Stop
                        Write-Verbose "[$operationId] Retrieved user object: $($SourceUserObj.DistinguishedName)"
                        # Add user to target domain group using the user object
                        Add-ADGroupMember -Identity $entry.TargetGroup -Members $SourceUserObj -Server $entry.TargetDomain -ErrorAction Stop -Verbose -WhatIf
                        $status = "Test Success"
                        $message = "Test mode: User would be successfully added to group"
                    }
                    else {
                        # Add user to group with retry logic
                        Invoke-WithRetry -ScriptBlock {
                            # Get the user object from source domain first
                            $SourceUserObj = Get-ADUser $entry.SourceUser -Server $entry.SourceDomain -Credential $ForeignAdminCreds -ErrorAction Stop
                            Write-Verbose "[$operationId] Retrieved user object: $($SourceUserObj.DistinguishedName)"
                            # Add user to target domain group using the user object
                            Add-ADGroupMember -Identity $entry.TargetGroup -Members $SourceUserObj -Server $entry.TargetDomain -ErrorAction Stop -Verbose
                            Write-Verbose "[$operationId] User $SourceDomainUser successfully added to $TargetDomainGroup"
                        } -MaxRetries $MaxRetries -DelaySeconds $RetryDelaySeconds -OperationName "Add user to group"

                        $status = "Success"
                        $message = "User successfully added to group"
                        $successCount++
                    }
                }
                catch {
                    $status = "Error"
                    $message = $_.Exception.Message
                    $errorCount++
                    # Log the error but don't re-throw it - just continue to next entry
                    Write-Warning "[$operationId] Error: $message"
                }

                $operationEndTime = Get-Date
                $operationDuration = ($operationEndTime - $operationStartTime).TotalSeconds

                # Prepare detailed log entry
                $logEntry = [PSCustomObject]@{
                    OperationId     = $operationId
                    Timestamp       = $operationStartTime.ToString("yyyy-MM-dd HH:mm:ss")
                    SourceUser      = $entry.SourceUser
                    SourceDomain    = $entry.SourceDomain
                    TargetGroup     = $entry.TargetGroup
                    TargetDomain    = $entry.TargetDomain
                    Status          = $status
                    Message         = $message
                    DurationSeconds = [math]::Round($operationDuration, 2)
                    ProcessedBy     = $env:USERNAME
                    ComputerName    = $env:COMPUTERNAME
                    TestMode        = $Test.IsPresent
                }

                $logEntries += $logEntry

                # Real-time logging for large operations
                if ($logEntries.Count % 10 -eq 0) {
                    Write-Verbose "Processed $($logEntries.Count) entries so far..."
                }
            }
        }

        $endTime = Get-Date
        $totalDuration = ($endTime - $startTime).TotalSeconds    # Add summary entry
        $summaryEntry = [PSCustomObject]@{
            OperationId     = "SUMMARY"
            Timestamp       = $endTime.ToString("yyyy-MM-dd HH:mm:ss")
            SourceUser      = "N/A"
            SourceDomain    = "N/A"
            TargetGroup     = "N/A"
            TargetDomain    = "N/A"
            Status          = "Summary"
            Message         = "Total: $processedCount, Success: $successCount, Errors: $errorCount, Skipped: $($processedCount - $successCount - $errorCount)"
            DurationSeconds = [math]::Round($totalDuration, 2)
            ProcessedBy     = $env:USERNAME
            ComputerName    = $env:COMPUTERNAME
            TestMode        = $Test.IsPresent
        }

        $logEntries += $summaryEntry

        # Log the results to CSV with error handling
        try {
            Write-Verbose "Logging $($logEntries.Count) results to $LogFilePath"
            $logEntries | Export-Csv -Path $LogFilePath -NoTypeInformation -Encoding UTF8
            Write-Verbose "Log file created successfully"
        }
        catch {
            Write-Error "Failed to write log file: $($_.Exception.Message)"
            # Fallback: try to write to temp location
            $fallbackLogPath = Join-Path $env:TEMP "ADUserToGroupLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
            $logEntries | Export-Csv -Path $fallbackLogPath -NoTypeInformation -Encoding UTF8
            Write-Warning "Log written to fallback location: $fallbackLogPath"
        }

        Write-Verbose "Operation completed. Total duration: $([math]::Round($totalDuration, 2)) seconds"
        Write-Progress -Activity "Adding Users to Groups" -Completed
    }
}

# MAIN EXECUTION

# Initialize error handling - use Continue to allow processing of other entries when errors occur
$ErrorActionPreference = "Continue"
$VerbosePreference = if ($PSBoundParameters['Verbose']) { "Continue" } else { "SilentlyContinue" }

Write-Verbose "Starting Add-ForeignUsertoFileShareADGroup script execution"
Write-Verbose "Parameters: InputCsvPath=$InputCsvPath, OutputFolderPath=$OutputFolderPath, Test=$Test"

# Validate prerequisites
try {
    # Temporarily set ErrorActionPreference to Stop for critical validation
    $ErrorActionPreference = "Stop"

    # Check if running on Windows
    if ($PSVersionTable.Platform -and $PSVersionTable.Platform -ne "Win32NT") {
        throw "This script requires Windows PowerShell on Windows Server/Desktop"
    }

    # Validate PowerShell version (5.1 minimum)
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        throw "This script requires PowerShell 5.1 or later. Current version: $($PSVersionTable.PSVersion)"
    }

    Write-Verbose "Prerequisites check passed - PowerShell $($PSVersionTable.PSVersion) on Windows"

    # Validate and import Active Directory module
    if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
        throw "ActiveDirectory PowerShell module is not available. Please install RSAT tools or run on a Domain Controller."
    }

    Import-Module ActiveDirectory -Force
    Write-Verbose "ActiveDirectory module imported successfully"

    # Validate CSV file and schema
    Write-Verbose "Validating CSV file: $InputCsvPath"
    if (-not (Test-CSVSchema -CsvPath $InputCsvPath)) {
        throw "CSV validation failed. Please check the file format and required columns."
    }

    # Create output folder if it doesn't exist
    if (-not (Test-Path $OutputFolderPath)) {
        New-Item -Path $OutputFolderPath -ItemType Directory -Force | Out-Null
        Write-Verbose "Created output folder: $OutputFolderPath"
    }
}
catch {
    Write-Error "Prerequisites validation failed: $($_.Exception.Message)"
    exit 1
}
finally {
    # Restore the original ErrorActionPreference for main processing
    $ErrorActionPreference = "Continue"
}

# Generate enhanced filenames with metadata
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$computerName = $env:COMPUTERNAME
$userName = $env:USERNAME
$mode = if ($Test) { "TEST" } else { "LIVE" }

$logFileName = "ADUserToGroupLog_${mode}_${timestamp}_${computerName}_${userName}.csv"
$transcriptFileName = "Transcript_${mode}_${timestamp}_${computerName}_${userName}.txt"
$summaryFileName = "Summary_${mode}_${timestamp}_${computerName}_${userName}.txt"

$logFilePath = Join-Path -Path $OutputFolderPath -ChildPath $logFileName
$transcriptFilePath = Join-Path -Path $OutputFolderPath -ChildPath $transcriptFileName
$summaryFilePath = Join-Path -Path $OutputFolderPath -ChildPath $summaryFileName

# Start enhanced transcript
try {
    Start-Transcript -Path $transcriptFilePath -Force
    Write-Host "=== Enhanced Add-ForeignUsertoFileShareADGroup Script ===" -ForegroundColor Green
    Write-Host "Execution started: $(Get-Date)" -ForegroundColor Green
    Write-Host "Mode: $mode" -ForegroundColor $(if ($Test) { "Yellow" } else { "Green" })
    Write-Host "Computer: $computerName" -ForegroundColor Green
    Write-Host "User: $userName" -ForegroundColor Green
    Write-Host "Input CSV: $InputCsvPath" -ForegroundColor Green
    Write-Host "Output Folder: $OutputFolderPath" -ForegroundColor Green
    Write-Host "Log File: $logFilePath" -ForegroundColor Green
    Write-Host "=========================================================" -ForegroundColor Green
}
catch {
    Write-Warning "Failed to start transcript: $($_.Exception.Message)"
}

try {
    # Set ErrorActionPreference to Stop for this critical section
    $ErrorActionPreference = "Stop"

    # Import and validate CSV data
    Write-Verbose "Importing CSV data from: $InputCsvPath"
    $csvData = Import-Csv -Path $InputCsvPath

    # Filter out invalid entries
    $validData = $csvData | Where-Object {
        -not ([string]::IsNullOrWhiteSpace($_.SourceDomain) -or
            [string]::IsNullOrWhiteSpace($_.SourceUser) -or
            [string]::IsNullOrWhiteSpace($_.TargetDomain) -or
            [string]::IsNullOrWhiteSpace($_.TargetGroup))
    }

    $invalidEntries = $csvData.Count - $validData.Count
    if ($invalidEntries -gt 0) {
        Write-Warning "Filtered out $invalidEntries invalid entries with missing data"
    }

    Write-Host "Processing $($validData.Count) valid user-group assignments" -ForegroundColor Green

    if ($validData.Count -eq 0) {
        throw "No valid data found in CSV file"
    }

    # Build enhanced object array from CSV data
    $userGroupData = @()
    foreach ($row in $validData) {
        $userGroupData += [PSCustomObject]@{
            SourceDomain = $row.SourceDomain.Trim()
            SourceUser   = $row.SourceUser.Trim()
            TargetDomain = $row.TargetDomain.Trim()
            TargetGroup  = $row.TargetGroup.Trim()
        }
    }

    # Get unique domains for connectivity pre-check
    $sourceDomains = $userGroupData | Select-Object -ExpandProperty SourceDomain -Unique
    $targetDomains = $userGroupData | Select-Object -ExpandProperty TargetDomain -Unique

    Write-Verbose "Source domains: $($sourceDomains -join ', ')"
    Write-Verbose "Target domains: $($targetDomains -join ', ')"

    # Execute the main operation
    $executionStart = Get-Date

    # Set ErrorActionPreference to Continue for main processing to handle individual errors gracefully
    $ErrorActionPreference = "Continue"

    # Ensure credentials are available before processing
    if (-not $ForeignAdminCreds) {
        Write-Host "Foreign domain credentials required for cross-domain operations..." -ForegroundColor Yellow
        $ForeignAdminCreds = Get-Credential -Message "Enter admin credentials for foreign domain access"
        if (-not $ForeignAdminCreds) {
            throw "Credentials are required for foreign domain access"
        }
    }

    # Process all user-group assignments
    $userGroupData | Add-ADUserToGroup -LogFilePath $logFilePath -Test:$Test -ForeignAdminCreds $ForeignAdminCreds -MaxRetries $MaxRetries -RetryDelaySeconds $RetryDelaySeconds

    $executionEnd = Get-Date
    $totalExecutionTime = ($executionEnd - $executionStart).TotalSeconds

    # Generate summary report
    if (Test-Path $logFilePath) {
        $logData = Import-Csv -Path $logFilePath
        $summaryData = $logData | Where-Object { $_.OperationId -eq "SUMMARY" } | Select-Object -First 1

        $summaryReport = @"
=== EXECUTION SUMMARY ===
Execution Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Mode: $mode
Computer: $computerName
User: $userName
Input File: $InputCsvPath
Output Folder: $OutputFolderPath

=== RESULTS ===
$($summaryData.Message)
Total Execution Time: $([math]::Round($totalExecutionTime, 2)) seconds

=== FILES GENERATED ===
Log File: $logFilePath
Transcript: $transcriptFilePath
Summary: $summaryFilePath

=== PERFORMANCE METRICS ===
Average Time per Operation: $([math]::Round($totalExecutionTime / $userGroupData.Count, 2)) seconds
Operations per Minute: $([math]::Round($userGroupData.Count / ($totalExecutionTime / 60), 2))
"@

        $summaryReport | Out-File -FilePath $summaryFilePath -Encoding UTF8
        Write-Host $summaryReport -ForegroundColor Green
    }

    Write-Host "Script execution completed successfully!" -ForegroundColor Green

    if ($Test) {
        Write-Host "TEST MODE: No actual changes were made. Review the log file for details on what would have been done." -ForegroundColor Yellow
    }
}
catch {
    Write-Host "Script execution failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Check the transcript file for detailed error information: $transcriptFilePath" -ForegroundColor Red

    # Log the error
    $errorLogPath = Join-Path -Path $OutputFolderPath -ChildPath "Error_${timestamp}.txt"
    $errorDetails = @"
Error occurred at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Error Message: $($_.Exception.Message)
Stack Trace: $($_.Exception.StackTrace)
Script Location: $($_.InvocationInfo.ScriptName):$($_.InvocationInfo.ScriptLineNumber)
"@
    $errorDetails | Out-File -FilePath $errorLogPath -Encoding UTF8

    exit 1
}
finally {
    # Stop transcript
    try {
        Stop-Transcript
    }
    catch {
        Write-Warning "Failed to stop transcript: $($_.Exception.Message)"
    }
    # ...existing code...
}