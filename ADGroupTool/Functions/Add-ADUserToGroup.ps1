function Add-ADUserToGroup {
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

    .PARAMETER ExponentialBackoff
    Whether to use exponential backoff for retry delays.

    .EXAMPLE
    $userGroupData | Add-ADUserToGroup -LogFilePath "C:\Logs\ADUserToGroupLog.csv" -Test -Verbose

    .NOTES
    Author: Steven Wight with GitHub Copilot
    Date: 23/01/2025
    Enhanced with comprehensive error handling and validation
    Compatible with: Windows Server 2012 R2+ and PowerShell 5.1+
    #>

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
        [ValidateRange(1, 10)]
        [int]$MaxRetries = 3,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 60)]
        [int]$RetryDelaySeconds = 5,

        [Parameter(Mandatory = $false)]
        [switch]$ExponentialBackoff
    )

    begin {
        $logEntries = @()
        $processedCount = 0
        $successCount = 0
        $errorCount = 0
        $alreadyMemberCount = 0
        $startTime = Get-Date
        $allUserGroupData = @()

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
        Write-Verbose "Total items to process: $totalItems"

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
                $sourceUserDN = ""

                try {
                    Write-Verbose "[$operationId] Retrieving user $($entry.SourceUser) from source domain $($entry.SourceDomain)"
                    Write-Verbose "[$operationId] Using credentials: $($ForeignAdminCreds.UserName)"

                    if ($Test) {
                        Write-Verbose "[$operationId] TEST MODE: Would add $SourceDomainUser to $TargetDomainGroup"

                        # Get the user object from source domain first
                        $SourceUserObj = Get-ADUser $entry.SourceUser -Server $entry.SourceDomain -Credential $ForeignAdminCreds -ErrorAction Stop
                        $sourceUserDN = $SourceUserObj.DistinguishedName
                        Write-Verbose "[$operationId] Retrieved user object: $sourceUserDN"

                        # Test adding user to target domain group using the user object
                        Add-ADGroupMember -Identity $entry.TargetGroup -Members $SourceUserObj -Server $entry.TargetDomain -ErrorAction Stop -Verbose -WhatIf
                        $status = "Test Success"
                        $message = "Test mode: User would be successfully added to group"
                    }
                    else {
                        # Add user to group with retry logic
                        Invoke-WithRetry -ScriptBlock {
                            # Get the user object from source domain first
                            $SourceUserObj = Get-ADUser $entry.SourceUser -Server $entry.SourceDomain -Credential $ForeignAdminCreds -ErrorAction Stop
                            $script:sourceUserDN = $SourceUserObj.DistinguishedName
                            Write-Verbose "[$operationId] Retrieved user object: $script:sourceUserDN"

                            # Add user to target domain group using the user object
                            Add-ADGroupMember -Identity $entry.TargetGroup -Members $SourceUserObj -Server $entry.TargetDomain -ErrorAction Stop -Verbose
                            Write-Verbose "[$operationId] User $SourceDomainUser successfully added to $TargetDomainGroup"
                        } -MaxRetries $MaxRetries -DelaySeconds $RetryDelaySeconds -OperationName "Add user to group" -ExponentialBackoff:$ExponentialBackoff

                        $sourceUserDN = $script:sourceUserDN
                        $status = "Success"
                        $message = "User successfully added to group"
                        $successCount++
                    }
                }
                catch {
                    $status = "Error"
                    $errorCode = if ($_.Exception.HResult) { "0x{0:X8}" -f $_.Exception.HResult } else { "Unknown" }

                    # Categorize common AD errors
                    $errorType = switch -Regex ($_.Exception.Message) {
                        "The server is not operational" { "ConnectivityError" }
                        "The user name or password is incorrect" { "AuthenticationError" }
                        "already a member of the group" {
                            $alreadyMemberCount++
                            $status = "AlreadyMember"
                            "AlreadyMember"
                        }
                        "Cannot find object|does not exist" { "ObjectNotFound" }
                        "Access.*denied" { "AccessDenied" }
                        default { "OtherError" }
                    }

                    $message = "$errorType ($errorCode): $($_.Exception.Message)"

                    # Don't increment error count for "already member" status
                    if ($status -ne "AlreadyMember") {
                        $errorCount++
                    }

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
                    SourceUserDN    = $sourceUserDN
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
        $totalDuration = ($endTime - $startTime).TotalSeconds

        # Add summary entry
        $summaryEntry = [PSCustomObject]@{
            OperationId     = "SUMMARY"
            Timestamp       = $endTime.ToString("yyyy-MM-dd HH:mm:ss")
            SourceUser      = "N/A"
            SourceDomain    = "N/A"
            SourceUserDN    = "N/A"
            TargetGroup     = "N/A"
            TargetDomain    = "N/A"
            Status          = "Summary"
            Message         = "Total: $processedCount, Success: $successCount, Errors: $errorCount, AlreadyMember: $alreadyMemberCount, Skipped: $($processedCount - $successCount - $errorCount - $alreadyMemberCount)"
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

        # Return summary information
        return [PSCustomObject]@{
            TotalProcessed = $processedCount
            SuccessCount = $successCount
            ErrorCount = $errorCount
            AlreadyMemberCount = $alreadyMemberCount
            SkippedCount = $processedCount - $successCount - $errorCount - $alreadyMemberCount
            TotalDurationSeconds = [math]::Round($totalDuration, 2)
            LogFilePath = $LogFilePath
        }
    }
}
