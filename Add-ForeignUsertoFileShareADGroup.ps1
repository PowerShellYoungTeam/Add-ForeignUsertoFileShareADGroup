param (
    [Parameter(Mandatory = $true)]
    [string]$InputCsvPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputFolderPath,

    [Parameter(Mandatory = $false)]
    [int]$TimeLimitInSeconds = 900
)

# Generate filenames for the log and transcript files
$timestamp = (Get-Date).ToString("yyyyMMdd")
$logFilePath = Join-Path -Path $OutputFolderPath -ChildPath "ADUserToGroupLog_$timestamp.csv"
$transcriptFilePath = Join-Path -Path $OutputFolderPath -ChildPath "Transcript_$timestamp.txt"

# Start transcript
Start-Transcript -Path $transcriptFilePath

# Import the CSV file
$data = Import-Csv -Path $InputCsvPath

# Extract the columns from the CSV
$sourceDomains = $data.SourceDomain
$sourceUsers = $data.SourceUser
$targetDomains = $data.TargetDomain
$targetGroups = $data.TargetGroup

# Call the Add-ADUserToGroup function with the extracted data
$sourceDomains, $sourceUsers, $targetDomains, $targetGroups | Add-ADUserToGroup -TimeLimitInSeconds $TimeLimitInSeconds -LogFilePath $logFilePath -Verbose

# Stop transcript
Stop-Transcript

function Add-ADUserToGroup {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string[]]$SourceDomain,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string[]]$SourceUser,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string[]]$TargetDomain,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string[]]$TargetGroup,

        [Parameter(Mandatory = $true)]
        [int]$TimeLimitInSeconds,

        [Parameter(Mandatory = $true)]
        [string]$LogFilePath
    )

    <#
    .SYNOPSIS
    Adds users from source domains to groups in target domains and logs the result.

    .DESCRIPTION
    This function adds specified users from source domains to specified groups in target domains.
    It waits for a specified time limit and then checks if the users are members of the groups.
    The result is logged to a CSV file.

    .PARAMETER SourceDomain
    The domains of the source users.

    .PARAMETER SourceUser
    The usernames of the source users.

    .PARAMETER TargetDomain
    The domains of the target groups.

    .PARAMETER TargetGroup
    The names of the target groups.

    .PARAMETER TimeLimitInSeconds
    The time limit to wait before checking if the users are members of the groups.

    .PARAMETER LogFilePath
    The file path where the log will be saved.

    .EXAMPLE
    "source.com", "source2.com" | Add-ADUserToGroup -SourceUser "user1", "user2" -TargetDomain "target.com", "target2.com" -TargetGroup "TargetGroup", "TargetGroup2" -TimeLimitInSeconds 60 -LogFilePath "C:\Logs\ADUserToGroupLog.csv" -Verbose

    .NOTES
    Author: Your Name
    Date: Today's Date
    #>

    begin {
        $logEntries = @()
    }

    process {
        for ($i = 0; $i -lt $SourceDomain.Length; $i++) {
            $currentSourceDomain = $SourceDomain[$i]
            $currentSourceUser = $SourceUser[$i]
            $currentTargetDomain = $TargetDomain[$i]
            $currentTargetGroup = $TargetGroup[$i]
            $SourceDomainUser = "$currentSourceDomain\$currentSourceUser"
            $TargetDomainGroup = "$currentTargetDomain\$currentTargetGroup"

            Write-Verbose "Adding $SourceDomainUser to $TargetDomainGroup"
            Write-Progress -Activity "Adding Users to Groups" -Status "Processing $currentSourceUser" -PercentComplete (($i / $SourceDomain.Length) * 100)

            try {
                # Add user to target domain group
                Add-ADGroupMember -Identity $TargetDomainGroup -Members $SourceDomainUser -ErrorAction Stop
                Write-Verbose "User $SourceDomainUser added to $TargetDomainGroup"

                # Wait for the specified time limit
                Write-Verbose "Waiting for $TimeLimitInSeconds seconds"
                Start-Sleep -Seconds $TimeLimitInSeconds

                # Confirm if the user is in the group
                Write-Verbose "Checking if $SourceDomainUser is a member of $TargetDomainGroup"
                $isMember = Get-ADGroupMember -Identity $TargetDomainGroup | Where-Object { $_.SamAccountName -eq $currentSourceUser }

                if ($isMember) {
                    $status = "Success"
                    $message = "$SourceDomainUser successfully added to $TargetDomainGroup"
                    Write-Verbose $message
                }
                else {
                    $status = "Failure"
                    $message = "$SourceDomainUser not found in $TargetDomainGroup after $TimeLimitInSeconds seconds"
                    Write-Verbose $message
                }
            }
            catch {
                $status = "Error"
                $message = $_.Exception.Message
                Write-Verbose "Error: $message"
            }

            # Prepare the log entry
            $logEntry = [PSCustomObject]@{
                Timestamp   = Get-Date
                UserName    = $currentSourceUser
                UserDomain  = $currentSourceDomain
                GroupName   = $currentTargetGroup
                GroupDomain = $currentTargetDomain
                Status      = $status
                Message     = $message
            }

            $logEntries += $logEntry
        }
    }

    end {
        # Log the results to CSV
        Write-Verbose "Logging results to $LogFilePath"
        $logEntries | Export-Csv -Path $LogFilePath -Append -NoTypeInformation
    }
}

# Example usage:
# "source.com", "source2.com" | Add-ADUserToGroup -SourceUser "user1", "user2" -TargetDomain "target.com", "target2.com" -TargetGroup "TargetGroup", "TargetGroup2" -TimeLimitInSeconds 60 -LogFilePath "C:\Logs\ADUserToGroupLog.csv" -Verbose