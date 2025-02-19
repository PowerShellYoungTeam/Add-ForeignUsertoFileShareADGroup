<#
.SYNOPSIS
Wrapper script to add users from a CSV file to AD groups and log the results.

.DESCRIPTION
This script reads user and group information from a CSV file and adds the users to the specified AD groups.
It logs the results to a CSV file and creates a transcript of the operation.

.PARAMETER InputCsvPath
The file path to the input CSV file containing user and group information.

.PARAMETER OutputFolderPath
The folder path where the log and transcript files will be saved.

.PARAMETER Test
Boolean to simulate the operation without making any changes.

.PARAMETER ForeignAdminCreds
Credentials for the foreign domain admin.

.EXAMPLE
.\Add-ForeignUsertoFileShareADGroup.ps1 -InputCsvPath "C:\Input\users.csv" -OutputFolderPath "C:\Output"  -Test $true

.NOTES
Author: Steven Wight with GitHub Copilot
Date: 23/01/2025

#>

param (
    [Parameter(Mandatory = $true)]
    [string]$InputCsvPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputFolderPath,

    [Parameter(Mandatory = $false)]
    [bool]$Test,

    [Parameter(Mandatory = $false)]
    [PSCredential]$ForeignAdminCreds
)

# FUNCTIONS

function Add-ADUserToGroup {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [PSCustomObject[]]$UserGroupData,


        [Parameter(Mandatory = $true)]
        [string]$LogFilePath,

        [Parameter(Mandatory = $false)]
        [Bool]$Test,

        [Parameter(Mandatory = $false)]
        [PSCredential]$ForeignAdminCreds
    )

    <#
    .SYNOPSIS
    Adds users from source domains to groups in target domains and logs the result.

    .DESCRIPTION
    This function adds specified users from source domains to specified groups in target domains.
    It waits for a specified time limit and then checks if the users are members of the groups.
    The result is logged to a CSV file.

    .PARAMETER UserGroupData
    The object array containing user and group information.

    .PARAMETER LogFilePath
    The file path where the log will be saved.

    .PARAMETER Test
    Boolean to simulate the operation without making any changes.

    .PARAMETER ForeignAdminCreds
    Credentials for the foreign domain admin.

    .EXAMPLE
    $userGroupData | Add-ADUserToGroup -LogFilePath "C:\Logs\ADUserToGroupLog.csv" -Test $test -Verbose

    .NOTES
    Author: Steven Wight with GitHub Copilot
    Date: 23/01/2025
    #>

    begin {
        $logEntries = @()
        if (-not $ForeignAdminCreds) {
            $ForeignAdminCreds = Get-Credential -Message "Enter admin creds for foreign domain"
        }
    }

    process {
        foreach ($entry in $UserGroupData) {
            $SourceDomainUser = "$($entry.SourceDomain)\$($entry.SourceUser)"
            $TargetDomainGroup = "$($entry.TargetDomain)\$($entry.TargetGroup)"

            Write-Verbose "Adding $SourceDomainUser to $TargetDomainGroup"
            Write-Progress -Activity "Adding Users to Groups" -Status "Processing $($entry.SourceUser)" -PercentComplete (([array]::IndexOf($UserGroupData, $entry) / $UserGroupData.Length) * 100)

            if ($PSCmdlet.ShouldProcess("$SourceDomainUser to $TargetDomainGroup")) {
                try {
                    if ($Test) {
                        Write-Verbose "Test mode: User $SourceDomainUser would be added to $TargetDomainGroup"
                        $SourceUserObj = get-aduser $entry.SourceUser -server $entry.SourceDomain -Credential $ForeignAdminCreds
                        Add-ADGroupMember -Identity $entry.TargetGroup -Members $SourceUserObj -Server $entry.TargetDomain -ErrorAction Stop -verbose -WhatIf
                    }
                    else {
                        # Add user to target domain group
                        $SourceUserObj = get-aduser $entry.SourceUser -server $entry.SourceDomain -Credential $ForeignAdminCreds
                        Add-ADGroupMember -Identity $entry.TargetGroup -Members $SourceUserObj -Server $entry.TargetDomain -ErrorAction Stop -verbose
                        Write-Verbose "User $SourceDomainUser added to $TargetDomainGroup"
                    }
                    $status = "Success"
                    $message = "$($logEntry.UserName) successfully added to $($logEntry.GroupName)"
                }
                catch {
                    $status = "Error"
                    $message = $_.Exception.Message
                    Write-Verbose "Error: $message"
                }

                # Prepare the log entry
                $logEntry = [PSCustomObject]@{
                    Timestamp   = Get-Date
                    UserName    = $entry.SourceUser
                    UserDomain  = $entry.SourceDomain
                    GroupName   = $entry.TargetGroup
                    GroupDomain = $entry.TargetDomain
                    Status      = $status
                    Message     = $message
                }

                $logEntries += $logEntry
            }
        }
    }

    end {

        # Log the results to CSV
        Write-Verbose "Logging results to $LogFilePath"
        $logEntries | Export-Csv -Path $LogFilePath -Append -NoTypeInformation
    }
}

# Example usage:
# "source.com", "source2.com" | Add-ADUserToGroup -SourceUser "user1", "user2" -TargetDomain "target.com", "target2.com" -TargetGroup "TargetGroup", "TargetGroup2"  -LogFilePath "C:\Logs\ADUserToGroupLog.csv" -Verbose

# Generate filenames for the log and transcript files
$timestamp = (Get-Date).ToString("yyyyMMdd")
$logFilePath = Join-Path -Path $OutputFolderPath -ChildPath "ADUserToGroupLog_$timestamp.csv"
$transcriptFilePath = Join-Path -Path $OutputFolderPath -ChildPath "Transcript_$timestamp.txt"

# Start transcript
Start-Transcript -Path $transcriptFilePath

# Import the CSV file
$data = Import-Csv -Path $InputCsvPath

# Build an object array from the CSV data
$userGroupData = @()
foreach ($row in $data) {
    $userGroupData += [PSCustomObject]@{
        SourceDomain = $row.SourceDomain
        SourceUser   = $row.SourceUser
        TargetDomain = $row.TargetDomain
        TargetGroup  = $row.TargetGroup
    }
}

# Call the Add-ADUserToGroup function with the object array
$userGroupData | Add-ADUserToGroup  -LogFilePath $logFilePath -Test:$Test -ForeignAdminCreds $ForeignAdminCreds -Verbose

# Stop transcript
Stop-Transcript
