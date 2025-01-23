function Add-ADUserToGroup {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$SourceDomain,

        [Parameter(Mandatory = $true)]
        [string]$SourceUser,

        [Parameter(Mandatory = $true)]
        [string]$TargetDomain,

        [Parameter(Mandatory = $true)]
        [string]$TargetGroup,

        [Parameter(Mandatory = $true)]
        [int]$TimeLimitInSeconds,

        [Parameter(Mandatory = $true)]
        [string]$LogFilePath
    )

    <#
    .SYNOPSIS
    Adds a user from a source domain to a group in a target domain and logs the result.

    .DESCRIPTION
    This function adds a specified user from a source domain to a specified group in a target domain.
    It waits for a specified time limit and then checks if the user is a member of the group.
    The result is logged to a CSV file.

    .PARAMETER SourceDomain
    The domain of the source user.

    .PARAMETER SourceUser
    The username of the source user.

    .PARAMETER TargetDomain
    The domain of the target group.

    .PARAMETER TargetGroup
    The name of the target group.

    .PARAMETER TimeLimitInSeconds
    The time limit to wait before checking if the user is a member of the group.

    .PARAMETER LogFilePath
    The file path where the log will be saved.

    .EXAMPLE
    Add-ADUserToGroup -SourceDomain "source.com" -SourceUser "user1" -TargetDomain "target.com" -TargetGroup "TargetGroup" -TimeLimitInSeconds 60 -LogFilePath "C:\Logs\ADUserToGroupLog.csv" -Verbose

    .NOTES
    Author: Your Name
    Date: Today's Date
    #>

    # Construct the full user and group names
    $SourceDomainUser = "$SourceDomain\$SourceUser"
    $TargetDomainGroup = "$TargetDomain\$TargetGroup"

    Write-Verbose "Adding $SourceDomainUser to $TargetDomainGroup"

    try {
        # Add user to target domain group
        Add-ADGroupMember -Identity $TargetDomainGroup -Members $SourceDomainUser -ErrorAction Stop
        Write-Verbose "User $SourceDomainUser added to $TargetDomainGroup"

        # Wait for the specified time limit
        Write-Verbose "Waiting for $TimeLimitInSeconds seconds"
        Start-Sleep -Seconds $TimeLimitInSeconds

        # Confirm if the user is in the group
        Write-Verbose "Checking if $SourceDomainUser is a member of $TargetDomainGroup"
        $isMember = Get-ADGroupMember -Identity $TargetDomainGroup | Where-Object { $_.SamAccountName -eq $SourceUser }

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

    # Log the result to CSV
    $logEntry = [PSCustomObject]@{
        Timestamp   = Get-Date
        UserName    = $SourceUser
        UserDomain  = $SourceDomain
        GroupName   = $TargetGroup
        GroupDomain = $TargetDomain
        Status      = $status
        Message     = $message
    }

    Write-Verbose "Logging result to $LogFilePath"
    $logEntry | Export-Csv -Path $LogFilePath -Append -NoTypeInformation
}

# Example usage:
# Add-ADUserToGroup -SourceDomain "source.com" -SourceUser "user1" -TargetDomain "target.com" -TargetGroup "TargetGroup" -TimeLimitInSeconds 60 -LogFilePath "C:\Logs\ADUserToGroupLog.csv" -Verbose