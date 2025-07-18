function Invoke-ADGroupOperation {
    <#
    .SYNOPSIS
    Main function to orchestrate adding users from a CSV file to AD groups with comprehensive validation and error handling.

    .DESCRIPTION
    This is the main entry point function that replicates the functionality of the original script.
    It reads user and group information from a CSV file and adds the users to the specified AD groups
    with comprehensive validation, retry logic, performance monitoring, and detailed logging.

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

    .PARAMETER ExponentialBackoff
    Whether to use exponential backoff for retry delays.

    .PARAMETER ValidateCredentials
    Whether to validate credentials against all domains before processing.

    .PARAMETER TestConnectivity
    Whether to test domain connectivity before processing.

    .EXAMPLE
    Invoke-ADGroupOperation -InputCsvPath "C:\Input\users.csv" -OutputFolderPath "C:\Output" -Test

    .EXAMPLE
    Invoke-ADGroupOperation -InputCsvPath "C:\Input\users.csv" -OutputFolderPath "C:\Output" -MaxRetries 5 -ValidateCredentials -TestConnectivity

    .NOTES
    Author: Steven Wight with GitHub Copilot
    Date: 18/07/2025
    Enhanced for modular use with comprehensive validation
    Compatible with: Windows Server 2012 R2+ and PowerShell 5.1+
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
        [int]$RetryDelaySeconds = 5,

        [Parameter(Mandatory = $false)]
        [switch]$ExponentialBackoff,

        [Parameter(Mandatory = $false)]
        [switch]$ValidateCredentials,

        [Parameter(Mandatory = $false)]
        [switch]$TestConnectivity
    )

    # Initialize error handling
    $ErrorActionPreference = "Continue"
    $VerbosePreference = if ($PSBoundParameters['Verbose']) { "Continue" } else { "SilentlyContinue" }

    Write-Verbose "Starting AD Group Operation"
    Write-Verbose "Parameters: InputCsvPath=$InputCsvPath, OutputFolderPath=$OutputFolderPath, Test=$Test"

    # Validate prerequisites
    try {
        $ErrorActionPreference = "Stop"

        # Check if running on Windows
        if ($PSVersionTable.Platform -and $PSVersionTable.Platform -ne "Win32NT") {
            throw "This function requires Windows PowerShell on Windows Server/Desktop"
        }

        # Validate PowerShell version (5.1 minimum)
        if ($PSVersionTable.PSVersion.Major -lt 5) {
            throw "This function requires PowerShell 5.1 or later. Current version: $($PSVersionTable.PSVersion)"
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
        $csvValidation = Test-CSVSchema -CsvPath $InputCsvPath
        if (-not $csvValidation.IsValid) {
            throw "CSV validation failed: $($csvValidation.Error)"
        }

        Write-Verbose "CSV validation passed: $($csvValidation.ValidRows)/$($csvValidation.TotalRows) valid rows"

        # Create output folder if it doesn't exist
        if (-not (Test-Path $OutputFolderPath)) {
            New-Item -Path $OutputFolderPath -ItemType Directory -Force | Out-Null
            Write-Verbose "Created output folder: $OutputFolderPath"
        }
    }
    catch {
        Write-Error "Prerequisites validation failed: $($_.Exception.Message)"
        return $false
    }
    finally {
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
        Write-Host "=== Enhanced AD Group Management Tool ===" -ForegroundColor Green
        Write-Host "Execution started: $(Get-Date)" -ForegroundColor Green
        Write-Host "Mode: $mode" -ForegroundColor $(if ($Test) { "Yellow" } else { "Green" })
        Write-Host "Computer: $computerName" -ForegroundColor Green
        Write-Host "User: $userName" -ForegroundColor Green
        Write-Host "Input CSV: $InputCsvPath" -ForegroundColor Green
        Write-Host "Output Folder: $OutputFolderPath" -ForegroundColor Green
        Write-Host "Log File: $logFilePath" -ForegroundColor Green
        Write-Host "=========================================" -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to start transcript: $($_.Exception.Message)"
    }

    try {
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

        # Get unique domains for connectivity and credential checks
        $sourceDomains = $userGroupData | Select-Object -ExpandProperty SourceDomain -Unique
        $targetDomains = $userGroupData | Select-Object -ExpandProperty TargetDomain -Unique
        $allDomains = @($sourceDomains) + @($targetDomains) | Select-Object -Unique

        Write-Verbose "Source domains: $($sourceDomains -join ', ')"
        Write-Verbose "Target domains: $($targetDomains -join ', ')"

        # Test domain connectivity if requested
        if ($TestConnectivity) {
            Write-Host "Testing domain connectivity..." -ForegroundColor Yellow
            $connectivityResults = Test-DomainConnectivity -Domains $allDomains

            if (-not $connectivityResults.Summary.AllReachable) {
                $unreachableDomains = $connectivityResults.UnreachableDomains -join ', '
                Write-Warning "Some domains are not reachable: $unreachableDomains"

                if (-not $Test) {
                    $choice = $host.UI.PromptForChoice(
                        "Connectivity Issues",
                        "Some domains are not reachable. Continue anyway?",
                        @("&Yes", "&No"),
                        1)
                    if ($choice -eq 1) {
                        throw "Operation aborted due to domain connectivity issues."
                    }
                }
            }
            else {
                Write-Host "All domains are reachable" -ForegroundColor Green
            }
        }

        # Ensure credentials are available before processing
        if (-not $ForeignAdminCreds) {
            Write-Host "Foreign domain credentials required for cross-domain operations..." -ForegroundColor Yellow
            $ForeignAdminCreds = Get-Credential -Message "Enter admin credentials for foreign domain access"
            if (-not $ForeignAdminCreds) {
                throw "Credentials are required for foreign domain access"
            }
        }

        # Validate credentials if requested
        if ($ValidateCredentials) {
            Write-Host "Validating credentials against domains..." -ForegroundColor Yellow
            $credentialValidationResults = @{}
            $allValid = $true

            foreach ($domain in $allDomains) {
                $validation = Test-Credential -Credential $ForeignAdminCreds -Domain $domain
                $credentialValidationResults[$domain] = $validation

                if (-not $validation.IsValid) {
                    Write-Warning "Credentials invalid for domain: $domain - $($validation.Error)"
                    $allValid = $false
                }
                else {
                    Write-Verbose "Credentials validated for domain: $domain"
                }
            }

            if (-not $allValid) {
                $invalidDomains = $credentialValidationResults.GetEnumerator() | Where-Object { -not $_.Value.IsValid } | ForEach-Object { $_.Key }

                if (-not $Test) {
                    $choice = $host.UI.PromptForChoice(
                        "Invalid Credentials",
                        "Credentials failed validation for domains: $($invalidDomains -join ', '). Continue anyway?",
                        @("&Yes", "&No"),
                        1)
                    if ($choice -eq 1) {
                        throw "Operation aborted due to invalid credentials."
                    }
                }
                else {
                    Write-Warning "Proceeding in test mode despite invalid credentials for: $($invalidDomains -join ', ')"
                }
            }
            else {
                Write-Host "All credential validations passed" -ForegroundColor Green
            }
        }

        # Execute the main operation
        $executionStart = Get-Date
        $ErrorActionPreference = "Continue"

        # Process all user-group assignments
        Write-Host "Starting user-group processing..." -ForegroundColor Green
        $results = $userGroupData | Add-ADUserToGroup -LogFilePath $logFilePath -Test:$Test -ForeignAdminCreds $ForeignAdminCreds -MaxRetries $MaxRetries -RetryDelaySeconds $RetryDelaySeconds -ExponentialBackoff:$ExponentialBackoff

        $executionEnd = Get-Date
        $totalExecutionTime = ($executionEnd - $executionStart).TotalSeconds

        # Generate summary report
        $summaryReport = @"
=== EXECUTION SUMMARY ===
Execution Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Mode: $mode
Computer: $computerName
User: $userName
Input File: $InputCsvPath
Output Folder: $OutputFolderPath

=== RESULTS ===
Total Processed: $($results.TotalProcessed)
Successful: $($results.SuccessCount)
Errors: $($results.ErrorCount)
Already Members: $($results.AlreadyMemberCount)
Skipped: $($results.SkippedCount)
Total Execution Time: $([math]::Round($totalExecutionTime, 2)) seconds

=== FILES GENERATED ===
Log File: $logFilePath
Transcript: $transcriptFilePath
Summary: $summaryFilePath

=== PERFORMANCE METRICS ===
Average Time per Operation: $([math]::Round($totalExecutionTime / $results.TotalProcessed, 2)) seconds
Operations per Minute: $([math]::Round($results.TotalProcessed / ($totalExecutionTime / 60), 2))
"@

        $summaryReport | Out-File -FilePath $summaryFilePath -Encoding UTF8
        Write-Host $summaryReport -ForegroundColor Green

        Write-Host "Operation completed successfully!" -ForegroundColor Green

        if ($Test) {
            Write-Host "TEST MODE: No actual changes were made. Review the log file for details on what would have been done." -ForegroundColor Yellow
        }

        return $true
    }
    catch {
        Write-Host "Operation failed: $($_.Exception.Message)" -ForegroundColor Red
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

        return $false
    }
    finally {
        # Stop transcript
        try {
            Stop-Transcript
        }
        catch {
            Write-Warning "Failed to stop transcript: $($_.Exception.Message)"
        }
    }
}
