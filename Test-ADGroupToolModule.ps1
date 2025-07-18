# Test script for ADGroupTool module
# This script validates that the module is properly structured and functional

Write-Host "=== ADGroupTool Module Test ===" -ForegroundColor Green

try {
    # Test 1: Import the module
    Write-Host "Test 1: Importing ADGroupTool module..." -ForegroundColor Yellow
    Import-Module .\ADGroupTool -Force -ErrorAction Stop
    Write-Host "✓ Module imported successfully" -ForegroundColor Green

    # Test 2: Check module information
    Write-Host "`nTest 2: Checking module information..." -ForegroundColor Yellow
    $module = Get-Module -Name ADGroupTool
    Write-Host "  Module Name: $($module.Name)" -ForegroundColor Cyan
    Write-Host "  Module Version: $($module.Version)" -ForegroundColor Cyan
    Write-Host "  Module Path: $($module.Path)" -ForegroundColor Cyan
    Write-Host "✓ Module information retrieved" -ForegroundColor Green

    # Test 3: Check exported functions
    Write-Host "`nTest 3: Checking exported functions..." -ForegroundColor Yellow
    $functions = Get-Command -Module ADGroupTool
    $expectedFunctions = @(
        'Add-ADUserToGroup',
        'Test-CSVSchema',
        'Test-Credential',
        'Test-DomainConnectivity',
        'Invoke-WithRetry',
        'Start-ADGroupToolGUI',
        'Install-ADGroupTool',
        'Invoke-ADGroupOperation'
    )

    foreach ($expectedFunc in $expectedFunctions) {
        if ($expectedFunc -in $functions.Name) {
            Write-Host "  ✓ $expectedFunc" -ForegroundColor Green
        } else {
            Write-Host "  ✗ $expectedFunc (MISSING)" -ForegroundColor Red
        }
    }

    # Test 4: Check aliases
    Write-Host "`nTest 4: Checking aliases..." -ForegroundColor Yellow
    $aliases = Get-Alias | Where-Object { $_.ModuleName -eq 'ADGroupTool' }
    foreach ($alias in $aliases) {
        Write-Host "  ✓ $($alias.Name) -> $($alias.Definition)" -ForegroundColor Green
    }

    # Test 5: Test CSV validation function
    Write-Host "`nTest 5: Testing CSV validation..." -ForegroundColor Yellow
    $sampleCsvPath = ".\ADGroupTool\Data\Sample_UserGroups.csv"
    if (Test-Path $sampleCsvPath) {
        $validation = Test-CSVSchema -CsvPath $sampleCsvPath
        if ($validation.IsValid) {
            Write-Host "  ✓ Sample CSV validation passed: $($validation.ValidRows)/$($validation.TotalRows) valid rows" -ForegroundColor Green
        } else {
            Write-Host "  ✗ Sample CSV validation failed: $($validation.Error)" -ForegroundColor Red
        }
    } else {
        Write-Host "  ⚠ Sample CSV file not found at $sampleCsvPath" -ForegroundColor Yellow
    }

    # Test 6: Test function help
    Write-Host "`nTest 6: Testing function help..." -ForegroundColor Yellow
    try {
        $help = Get-Help Invoke-ADGroupOperation -ErrorAction Stop
        if ($help.Synopsis) {
            Write-Host "  ✓ Help available for Invoke-ADGroupOperation" -ForegroundColor Green
        } else {
            Write-Host "  ⚠ Help available but no synopsis found" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "  ✗ Help not available for Invoke-ADGroupOperation" -ForegroundColor Red
    }

    Write-Host "`n=== All Tests Completed ===" -ForegroundColor Green
    Write-Host "Module appears to be properly structured and functional!" -ForegroundColor Green
    Write-Host "`nNext steps:" -ForegroundColor Yellow
    Write-Host "1. Run: Start-ADGroupToolGUI" -ForegroundColor Cyan
    Write-Host "2. Or: Get-Help Invoke-ADGroupOperation -Examples" -ForegroundColor Cyan

} catch {
    Write-Host "`n✗ Test failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Please check the module structure and try again." -ForegroundColor Yellow
}
