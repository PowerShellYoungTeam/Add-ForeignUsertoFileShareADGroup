function Test-DomainConnectivity {
    <#
    .SYNOPSIS
    Tests network connectivity and domain controller availability for specified domains.

    .DESCRIPTION
    This function tests both network connectivity and domain controller availability
    for a list of domains. It verifies that domains are reachable and that domain
    controllers can be contacted for AD operations.

    .PARAMETER Domains
    Array of domain names to test connectivity for.

    .PARAMETER TimeoutSeconds
    Timeout in seconds for connectivity tests. Default is 10 seconds.

    .EXAMPLE
    Test-DomainConnectivity -Domains @("contoso.com", "fabrikam.com")

    .NOTES
    Author: Steven Wight with GitHub Copilot
    Compatible with: Windows Server 2012 R2+ and PowerShell 5.1+
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$Domains,

        [Parameter(Mandatory = $false)]
        [int]$TimeoutSeconds = 10
    )

    $results = @{}

    foreach ($domain in $Domains) {
        Write-Verbose "Testing connectivity to domain: $domain"

        try {
            # Test basic network connectivity
            $pingResult = Test-Connection -ComputerName $domain -Count 1 -Quiet -ErrorAction SilentlyContinue

            if ($pingResult) {
                Write-Verbose "Basic connectivity to ${domain}: Success"

                # Try to query domain controller
                $dcFound = $false
                $dcError = $null

                try {
                    # Attempt to get domain information
                    Add-Type -AssemblyName System.DirectoryServices
                    $directoryContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $domain)
                    $domainObj = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($directoryContext)

                    if ($domainObj -and $domainObj.DomainControllers.Count -gt 0) {
                        $dcFound = $true
                        $dcName = $domainObj.DomainControllers[0].Name
                        Write-Verbose "Domain controller found for ${domain}: $dcName"
                    }
                }
                catch {
                    $dcError = $_.Exception.Message
                    Write-Verbose "Could not retrieve domain controller for ${domain}: $dcError"
                }

                $results[$domain] = [PSCustomObject]@{
                    Domain = $domain
                    Reachable = $true
                    DCFound = $dcFound
                    DCName = if ($dcFound) { $dcName } else { $null }
                    Error = $dcError
                    TestTime = Get-Date
                }
            }
            else {
                Write-Verbose "Basic connectivity to ${domain}: Failed"
                $results[$domain] = [PSCustomObject]@{
                    Domain = $domain
                    Reachable = $false
                    DCFound = $false
                    DCName = $null
                    Error = "Domain not reachable via ping"
                    TestTime = Get-Date
                }
            }
        }
        catch {
            $errorMessage = $_.Exception.Message
            Write-Verbose "Error testing connectivity to ${domain}: $errorMessage"

            $results[$domain] = [PSCustomObject]@{
                Domain = $domain
                Reachable = $false
                DCFound = $false
                DCName = $null
                Error = $errorMessage
                TestTime = Get-Date
            }
        }
    }

    # Summary information
    $totalDomains = $Domains.Count
    $reachableDomains = ($results.Values | Where-Object { $_.Reachable }).Count
    $domainsWithDC = ($results.Values | Where-Object { $_.DCFound }).Count

    Write-Verbose "Domain connectivity summary: $reachableDomains/$totalDomains reachable, $domainsWithDC/$totalDomains with accessible DCs"

    # Return results with summary
    return [PSCustomObject]@{
        Summary = [PSCustomObject]@{
            TotalDomains = $totalDomains
            ReachableDomains = $reachableDomains
            DomainsWithDC = $domainsWithDC
            AllReachable = ($reachableDomains -eq $totalDomains)
            AllDCsAccessible = ($domainsWithDC -eq $totalDomains)
        }
        Results = $results
        UnreachableDomains = ($results.Values | Where-Object { -not $_.Reachable }).Domain
        DomainsWithoutDC = ($results.Values | Where-Object { -not $_.DCFound }).Domain
    }
}
