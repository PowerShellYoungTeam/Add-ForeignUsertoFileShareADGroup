# Domain Connectivity Diagnostic Script
# Use this script to diagnose domain connectivity issues before running the main script

param(
    [Parameter(Mandatory = $false)]
    [string[]]$Domains = @("SANHQ", "SANUK"),

    [Parameter(Mandatory = $false)]
    [PSCredential]$Credential
)

Write-Host "=== Domain Connectivity Diagnostic Tool ===" -ForegroundColor Green
Write-Host ""

function Test-DomainBasics {
    param([string]$DomainName)

    Write-Host "Testing domain: $DomainName" -ForegroundColor Yellow

    # Test DNS resolution
    try {
        $dnsResult = Resolve-DnsName -Name $DomainName -Type A -ErrorAction Stop
        Write-Host "  ✓ DNS Resolution: Success" -ForegroundColor Green
        Write-Host "    IP Addresses: $($dnsResult.IPAddress -join ', ')" -ForegroundColor Gray
    }
    catch {
        Write-Host "  ✗ DNS Resolution: Failed - $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }

    # Test basic network connectivity
    try {
        $pingResult = Test-Connection -ComputerName $DomainName -Count 1 -Quiet -ErrorAction Stop
        if ($pingResult) {
            Write-Host "  ✓ Network Ping: Success" -ForegroundColor Green
        }
        else {
            Write-Host "  ✗ Network Ping: Failed" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "  ✗ Network Ping: Failed - $($_.Exception.Message)" -ForegroundColor Red
    }

    # Test AD connectivity
    try {
        $dcResult = Get-ADDomainController -Domain $DomainName -Credential $Credential -ErrorAction Stop
        Write-Host "  ✓ AD Connectivity: Success" -ForegroundColor Green
        Write-Host "    Domain Controller: $($dcResult.Name)" -ForegroundColor Gray
        Write-Host "    Domain: $($dcResult.Domain)" -ForegroundColor Gray
        Write-Host "    Forest: $($dcResult.Forest)" -ForegroundColor Gray
        return $true
    }
    catch [Microsoft.ActiveDirectory.Management.ADServerDownException] {
        Write-Host "  ✗ AD Connectivity: Domain controller unavailable" -ForegroundColor Red
        Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor Red
    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        Write-Host "  ✗ AD Connectivity: Domain not found or not accessible" -ForegroundColor Red
        Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor Red
    }
    catch [System.Security.Authentication.AuthenticationException] {
        Write-Host "  ✗ AD Connectivity: Authentication failed" -ForegroundColor Red
        Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor Red
    }
    catch {
        Write-Host "  ✗ AD Connectivity: Failed" -ForegroundColor Red
        Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "    Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
    }

    return $false
}

# Check if running as admin
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
Write-Host "Running as Administrator: $isAdmin" -ForegroundColor $(if ($isAdmin) { "Green" } else { "Yellow" })

# Check AD module
$adModule = Get-Module -Name ActiveDirectory -ListAvailable
if ($adModule) {
    Write-Host "ActiveDirectory Module: Available (Version: $($adModule.Version))" -ForegroundColor Green
}
else {
    Write-Host "ActiveDirectory Module: Not Available - Install RSAT Tools" -ForegroundColor Red
    exit 1
}

Write-Host ""

# Get credentials if not provided
if (-not $Credential) {
    Write-Host "No credentials provided - testing with current user context" -ForegroundColor Yellow
    Write-Host "For foreign domains, you may need to provide credentials." -ForegroundColor Yellow

    $response = Read-Host "Do you want to provide domain credentials? (y/n)"
    if ($response -eq 'y' -or $response -eq 'Y') {
        try {
            $Credential = Get-Credential -Message "Enter domain credentials"
        }
        catch {
            Write-Host "Failed to get credentials, continuing with current user" -ForegroundColor Yellow
        }
    }
}

if ($Credential) {
    Write-Host "Testing with credentials for: $($Credential.UserName)" -ForegroundColor Green
}
else {
    Write-Host "Testing with current user: $($env:USERNAME)@$($env:USERDOMAIN)" -ForegroundColor Green
}

Write-Host ""

# Test each domain
$results = @()
foreach ($domain in $Domains) {
    $success = Test-DomainBasics -DomainName $domain
    $results += [PSCustomObject]@{
        Domain  = $domain
        Success = $success
    }
    Write-Host ""
}

# Summary
Write-Host "=== Summary ===" -ForegroundColor Green
$results | Format-Table -AutoSize

$successCount = ($results | Where-Object { $_.Success }).Count
$totalCount = $results.Count

Write-Host "Successful connections: $successCount/$totalCount" -ForegroundColor $(if ($successCount -eq $totalCount) { "Green" } else { "Yellow" })

if ($successCount -lt $totalCount) {
    Write-Host ""
    Write-Host "Troubleshooting Tips:" -ForegroundColor Yellow
    Write-Host "1. Check network connectivity and firewalls" -ForegroundColor White
    Write-Host "2. Verify DNS settings can resolve domain names" -ForegroundColor White
    Write-Host "3. Ensure proper domain credentials are provided" -ForegroundColor White
    Write-Host "4. Check if RSAT tools are properly installed" -ForegroundColor White
    Write-Host "5. Verify domain trust relationships" -ForegroundColor White
    Write-Host "6. Use -SkipDomainConnectivityCheck in main script if needed" -ForegroundColor White
}

Write-Host ""
Write-Host "You can run the main script with -SkipDomainConnectivityCheck to bypass connectivity validation" -ForegroundColor Cyan
