# Test script for OFlights function - local testing without timer trigger
# This bypasses the timer trigger and tests the core logic directly

$CAPWATCHDATADIR = "$($env:HOME)\data\CAPWatch"
Push-Location $CAPWATCHDATADIR

# Include shared Functions
. "$PSScriptRoot\shared\shared.ps1"

Write-Host "Testing OFlights function..."
Write-Host "CAPWATCHDATADIR: $CAPWATCHDATADIR"
Write-Host ""

# Import the O-Flight CSV file
$allFlightRecords = @()
try {
    $allFlightRecords = Import-Csv "$($CAPWATCHDATADIR)\OFlight.txt" -ErrorAction Stop
    Write-Host "✓ Successfully imported $($allFlightRecords.Count) flight records from OFlight.txt"
} catch {
    Write-Host "✗ Failed to import OFlight.txt. Error: $($_.Exception.Message)"
    exit 1
}

# Create mock users for testing (replace with GetAllUsers when Azure auth is available)
Write-Host "✓ Creating mock user list for testing..."
$allUsers = @(
    @{ employeeId = "139034" },
    @{ employeeId = "173378" },
    @{ employeeId = "322448" },
    @{ employeeId = "331052" },
    @{ employeeId = "350885" },
    @{ employeeId = "368146" }
)

# Create a hash set of valid employee IDs for quick lookup
$validEmployeeIds = @{}
foreach ($user in $allUsers) {
    if ($user.employeeId) {
        $validEmployeeIds[$user.employeeId] = $true
    }
}
Write-Host "✓ Built hash table with $($validEmployeeIds.Count) valid employee IDs"
Write-Host ""

# Filter and process flight data
Write-Host "Processing flight data..."
$syllabusData = $allFlightRecords | 
    Where-Object { $_.Syllabus -in @("6", "7", "8", "9", "10") } |
    Group-Object -Property CAPID, Syllabus |
    ForEach-Object {
        $parts = $_.Name -split ', '
        [PSCustomObject]@{
            CAPID = $parts[0]
            Syllabus = $parts[1]
            FirstFlight = ($_.Group | Sort-Object FltDate | Select-Object -First 1).FltDate
        }
    } |
    Where-Object { $validEmployeeIds.ContainsKey($_.CAPID) }

Write-Host "✓ Processed flight data: Found $($syllabusData.Count) valid O-Flight records"
Write-Host ""

# Display sample data
Write-Host "Sample O-Flight records (first 10):"
$syllabusData | Select-Object -First 10 | Format-Table -AutoSize

Write-Host ""
Write-Host "✓ OFlights test completed successfully!"
Write-Host ""
Write-Host "To test Cosmos DB sync, you'll need Azure authentication and the following environment variables:"
Write-Host "  - CosmosDbConnectionString"
Write-Host "  - CosmosDbDatabase (should be: orientation-flights)"
Write-Host "  - CosmosDbContainer (should be: syllabus)"
