function Test-CSVSchema {
    <#
    .SYNOPSIS
    Validates the schema of a CSV file for AD Group Management operations.

    .DESCRIPTION
    This function validates that a CSV file contains the required columns for adding users to AD groups.
    It checks for required columns and identifies any empty rows that would be skipped during processing.

    .PARAMETER CsvPath
    The file path to the CSV file to validate.

    .EXAMPLE
    Test-CSVSchema -CsvPath "C:\Input\users.csv"

    .NOTES
    Author: Steven Wight with GitHub Copilot
    Compatible with: Windows Server 2012 R2+ and PowerShell 5.1+
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path $_ -PathType Leaf })]
        [string]$CsvPath
    )

    $requiredColumns = @('SourceDomain', 'SourceUser', 'TargetDomain', 'TargetGroup')

    try {
        Write-Verbose "Validating CSV schema for file: $CsvPath"
        $csvData = Import-Csv -Path $CsvPath -ErrorAction Stop

        if ($csvData.Count -eq 0) {
            throw "CSV file is empty or contains no data rows"
        }

        $csvColumns = $csvData[0].PSObject.Properties.Name
        Write-Verbose "Found columns: $($csvColumns -join ', ')"

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
            Write-Warning "Found $($emptyRows.Count) rows with missing data. These will be skipped during processing."
        }

        $validRows = $csvData.Count - $emptyRows.Count
        Write-Verbose "CSV schema validation passed. Found $($csvData.Count) total rows, $validRows valid rows."

        return [PSCustomObject]@{
            IsValid = $true
            TotalRows = $csvData.Count
            ValidRows = $validRows
            EmptyRows = $emptyRows.Count
            Columns = $csvColumns
            MissingColumns = @()
        }
    }
    catch {
        Write-Error "CSV validation failed: $($_.Exception.Message)"
        return [PSCustomObject]@{
            IsValid = $false
            TotalRows = 0
            ValidRows = 0
            EmptyRows = 0
            Columns = @()
            MissingColumns = $missingColumns
            Error = $_.Exception.Message
        }
    }
}
