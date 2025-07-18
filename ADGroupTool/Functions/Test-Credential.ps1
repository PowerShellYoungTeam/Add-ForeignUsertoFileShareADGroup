function Test-Credential {
    <#
    .SYNOPSIS
    Validates credentials against a specified domain.

    .DESCRIPTION
    This function validates provided credentials against a domain controller to ensure
    they are valid before attempting AD operations. Uses DirectoryServices for validation.

    .PARAMETER Credential
    The PSCredential object to validate.

    .PARAMETER Domain
    The domain name to validate the credentials against.

    .EXAMPLE
    Test-Credential -Credential $creds -Domain "contoso.com"

    .NOTES
    Author: Steven Wight with GitHub Copilot
    Compatible with: Windows Server 2012 R2+ and PowerShell 5.1+
    Requires: System.DirectoryServices.AccountManagement
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSCredential]$Credential,

        [Parameter(Mandatory = $true)]
        [string]$Domain
    )

    try {
        Write-Verbose "Validating credentials for user '$($Credential.UserName)' against domain '$Domain'"

        # Load required assembly
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement

        $contextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
        $principalContext = $null

        try {
            $principalContext = New-Object System.DirectoryServices.AccountManagement.PrincipalContext $contextType, $Domain

            # Extract username without domain prefix if present
            $username = $Credential.UserName
            if ($username.Contains('\')) {
                $username = $username.Split('\')[1]
            } elseif ($username.Contains('@')) {
                $username = $username.Split('@')[0]
            }

            $password = $Credential.GetNetworkCredential().Password
            $valid = $principalContext.ValidateCredentials($username, $password)

            if ($valid) {
                Write-Verbose "Credentials validated successfully for domain: $Domain"
                return [PSCustomObject]@{
                    IsValid = $true
                    Domain = $Domain
                    Username = $Credential.UserName
                    ValidationTime = Get-Date
                    Error = $null
                }
            }
            else {
                Write-Warning "Invalid credentials for domain: $Domain"
                return [PSCustomObject]@{
                    IsValid = $false
                    Domain = $Domain
                    Username = $Credential.UserName
                    ValidationTime = Get-Date
                    Error = "Authentication failed - invalid username or password"
                }
            }
        }
        finally {
            if ($principalContext) {
                $principalContext.Dispose()
            }
        }
    }
    catch {
        $errorMessage = "Error validating credentials for domain $Domain`: $($_.Exception.Message)"
        Write-Warning $errorMessage

        return [PSCustomObject]@{
            IsValid = $false
            Domain = $Domain
            Username = $Credential.UserName
            ValidationTime = Get-Date
            Error = $errorMessage
        }
    }
}
