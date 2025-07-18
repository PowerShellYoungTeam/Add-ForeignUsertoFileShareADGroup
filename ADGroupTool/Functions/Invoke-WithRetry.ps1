function Invoke-WithRetry {
    <#
    .SYNOPSIS
    Executes a script block with configurable retry logic and exponential backoff.

    .DESCRIPTION
    This function executes a script block and automatically retries on failure with
    configurable retry attempts and delays. Supports exponential backoff for certain
    error types and provides detailed logging of retry attempts.

    .PARAMETER ScriptBlock
    The script block to execute with retry logic.

    .PARAMETER MaxRetries
    Maximum number of retry attempts. Default is 3.

    .PARAMETER DelaySeconds
    Initial delay in seconds between retry attempts. Default is 5.

    .PARAMETER OperationName
    Descriptive name for the operation being performed (for logging).

    .PARAMETER ExponentialBackoff
    Whether to use exponential backoff for retry delays. Default is $false.

    .PARAMETER RetryableErrors
    Array of error patterns that should trigger a retry. If not specified, all errors trigger retries.

    .EXAMPLE
    Invoke-WithRetry -ScriptBlock { Get-ADUser "testuser" } -MaxRetries 3 -DelaySeconds 5 -OperationName "Get AD User"

    .NOTES
    Author: Steven Wight with GitHub Copilot
    Compatible with: Windows Server 2012 R2+ and PowerShell 5.1+
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 10)]
        [int]$MaxRetries = 3,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 300)]
        [int]$DelaySeconds = 5,

        [Parameter(Mandatory = $false)]
        [string]$OperationName = "Operation",

        [Parameter(Mandatory = $false)]
        [switch]$ExponentialBackoff,

        [Parameter(Mandatory = $false)]
        [string[]]$RetryableErrors = @()
    )

    $attempt = 1
    $currentDelay = $DelaySeconds

    do {
        try {
            Write-Verbose "Attempting $OperationName (Attempt $attempt of $($MaxRetries + 1))"
            $result = & $ScriptBlock
            Write-Verbose "$OperationName succeeded on attempt $attempt"
            return $result
        }
        catch {
            $errorMessage = $_.Exception.Message
            $shouldRetry = $true

            # Check if this is a retryable error (if specific patterns are defined)
            if ($RetryableErrors.Count -gt 0) {
                $shouldRetry = $false
                foreach ($pattern in $RetryableErrors) {
                    if ($errorMessage -match $pattern) {
                        $shouldRetry = $true
                        break
                    }
                }
            }

            # Categorize error for logging
            $errorCategory = "Unknown"
            switch -Regex ($errorMessage) {
                "server is not operational|RPC server" { $errorCategory = "Connectivity" }
                "access.*denied|unauthorized" { $errorCategory = "Authentication" }
                "timeout|timed out" { $errorCategory = "Timeout" }
                "not found|does not exist" { $errorCategory = "NotFound" }
                "already.*member|already exists" { $errorCategory = "AlreadyExists" }
                default { $errorCategory = "Other" }
            }

            Write-Warning "$OperationName failed on attempt $attempt (Category: $errorCategory): $errorMessage"

            if ($shouldRetry -and $attempt -le $MaxRetries) {
                Write-Verbose "Waiting $currentDelay seconds before retry..."
                Start-Sleep -Seconds $currentDelay

                # Apply exponential backoff if enabled
                if ($ExponentialBackoff) {
                    $currentDelay = [math]::Min($currentDelay * 2, 300) # Cap at 5 minutes
                }

                $attempt++
            }
            else {
                if (-not $shouldRetry) {
                    Write-Error "$OperationName failed with non-retryable error: $errorMessage"
                }
                else {
                    Write-Error "$OperationName failed after $($MaxRetries + 1) attempts"
                }
                throw
            }
        }
    } while ($attempt -le ($MaxRetries + 1))
}
