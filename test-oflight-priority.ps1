# Test script for OFlight Priority calculation - local testing
# This script runs the prioritization algorithm and optionally creates a flight schedule

Write-Host "==================== OFlight Priority Test Script ====================" -ForegroundColor Cyan
Write-Host ""

# Set default paths
$memberPath = "$env:HOME\data\CAPWatch\Member.txt"
$oflightPath = "$env:HOME\data\CAPWatch\OFlight.txt"
$outputPath = "$PSScriptRoot\OFlightPriority.csv"
$schedulePath = "$PSScriptRoot\OFlightSchedule.csv"

# Check if data files exist
Write-Host "Checking data files..." -ForegroundColor Yellow
if (-not (Test-Path $memberPath)) {
    Write-Host "❌ Error: Member.txt not found at $memberPath" -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $oflightPath)) {
    Write-Host "❌ Error: OFlight.txt not found at $oflightPath" -ForegroundColor Red
    exit 1
}
Write-Host "✅ Data files found" -ForegroundColor Green
Write-Host ""

# Run the prioritization script
Write-Host "Running prioritization algorithm..." -ForegroundColor Yellow
Write-Host "================================================================" -ForegroundColor Cyan

# Example 1: Just generate prioritized list (no schedule)
& "$PSScriptRoot\Get-OFlightPriority.ps1" `
    -MemberPath $memberPath `
    -OFlightsPath $oflightPath `
    -OutputCsv $outputPath

Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

# Ask if user wants to save to Cosmos DB
Write-Host "Would you like to save results to Cosmos DB? (Y/N)" -ForegroundColor Yellow
$saveToCosmosDb = Read-Host
$cosmosDbSwitch = @{}
if ($saveToCosmosDb -eq "Y" -or $saveToCosmosDb -eq "y") {
    $cosmosDbSwitch = @{ SaveToCosmosDb = $true }
}
Write-Host ""

# Ask if user wants to create a schedule
Write-Host "Would you like to create a flight schedule? (Y/N)" -ForegroundColor Yellow
$createSchedule = Read-Host

if ($createSchedule -eq "Y" -or $createSchedule -eq "y") {
    Write-Host ""
    Write-Host "How many flight slots are available?" -ForegroundColor Yellow
    $slots = Read-Host

    Write-Host "Maximum flights per squadron? (press Enter for no limit)" -ForegroundColor Yellow
    $maxPerSquadron = Read-Host

    Write-Host ""
    Write-Host "Creating flight schedule..." -ForegroundColor Yellow
    Write-Host "================================================================" -ForegroundColor Cyan

    if ($maxPerSquadron) {
        & "$PSScriptRoot\Get-OFlightPriority.ps1" `
            -MemberPath $memberPath `
            -OFlightsPath $oflightPath `
            -OutputCsv $outputPath `
            -OutputScheduleCsv $schedulePath `
            -TotalSlots ([int]$slots) `
            -MaxPerSquadron ([int]$maxPerSquadron) `
            @cosmosDbSwitch
    } else {
        & "$PSScriptRoot\Get-OFlightPriority.ps1" `
            -MemberPath $memberPath `
            -OFlightsPath $oflightPath `
            -OutputCsv $outputPath `
            -OutputScheduleCsv $schedulePath `
            -TotalSlots ([int]$slots) `
            @cosmosDbSwitch
    }

    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
} else {
    # No schedule, just priority list
    if ($cosmosDbSwitch.Count -gt 0) {
        & "$PSScriptRoot\Get-OFlightPriority.ps1" `
            -MemberPath $memberPath `
            -OFlightsPath $oflightPath `
            -OutputCsv $outputPath `
            @cosmosDbSwitch
    }
}

Write-Host "✅ OFlight Priority test completed!" -ForegroundColor Green
Write-Host ""
Write-Host "Output files:" -ForegroundColor Cyan
Write-Host "  - Full Priority List: $outputPath" -ForegroundColor White
if (Test-Path $schedulePath) {
    Write-Host "  - Flight Schedule: $schedulePath" -ForegroundColor White
}
Write-Host ""
