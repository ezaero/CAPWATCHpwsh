<#
.SYNOPSIS
    Processes O-Flight data from CAPWATCH and syncs it to Cosmos DB.

.DESCRIPTION
    This script performs the following tasks:
    1. Imports O-Flight records from the OFlight.txt file
    2. Filters flights for specific syllabuses (6, 7, 8, 9, 10)
    3. Groups flights by CAPID and Syllabus to identify first flight dates
    4. Validates that CAPIDs exist in the current Azure AD user list
    5. Syncs flight data to Cosmos DB, only updating when data has changed
    6. Logs all actions for auditing purposes

.PARAMETER None
    This function uses environment variables for configuration

.NOTES
    - Requires $CAPWATCHDATADIR environment variable to be set
    - Requires Cosmos DB configuration in environment variables
    - Requires Microsoft Graph connection to be established
#>

param($Timer)

# Set working directory to folder with all CAPWATCH CSV Text Files
$CAPWATCHDATADIR = "$($env:HOME)\data\CAPWatch"
Push-Location $CAPWATCHDATADIR

# Include shared Functions
. "$PSScriptRoot\..\shared\shared.ps1"

# Authenticate to Microsoft Graph using managed identity
try {
    $MSGraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token
    Connect-MgGraph -AccessToken $MSGraphAccessToken -NoWelcome
    Write-Log "Successfully authenticated to Microsoft Graph"
} catch {
    Write-Log "Failed to authenticate to Microsoft Graph. Error: $($_.Exception.Message)"
    exit 1
}

# Import the O-Flight CSV file
$allFlightRecords = @()
try {
    $allFlightRecords = Import-Csv "$($CAPWATCHDATADIR)\OFlight.txt" -ErrorAction Stop
    Write-Log "Successfully imported $($allFlightRecords.Count) flight records from OFlight.txt"
} catch {
    Write-Log "Failed to import OFlight.txt. Error: $($_.Exception.Message)"
    exit 1
}

# Get all users to validate CAPIDs
$allUsers = GetAllUsers
Write-Log "Retrieved $($allUsers.Count) users from Azure AD"

# Create a hash set of valid employee IDs for quick lookup
$validEmployeeIds = @{}
foreach ($user in $allUsers) {
    if ($user.employeeId) {
        $validEmployeeIds[$user.employeeId] = $true
    }
}
Write-Log "Built hash table with $($validEmployeeIds.Count) valid employee IDs"

# Filter and process flight data
# 1. Filter for specific syllabuses (6, 7, 8, 9, 10)
# 2. Group by CAPID and Syllabus to get first flight
# 3. Filter out invalid CAPIDs
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

Write-Log "Processed flight data: Found $($syllabusData.Count) valid O-Flight records"

# Sync to Cosmos DB
if ($syllabusData.Count -eq 0) {
    Write-Log "No valid flight records to sync. Exiting."
    exit 0
}

try {
    $cosmosConfig = Get-CosmosDbConnection
    
    if (-not $cosmosConfig.ConnectionString -or -not $cosmosConfig.Database -or -not $cosmosConfig.Container) {
        Write-Log "Error: Cosmos DB configuration incomplete. Verify CosmosDbConnectionString, CosmosDbDatabase, and CosmosDbContainer environment variables."
        exit 1
    }

    Write-Log "Syncing $($syllabusData.Count) flight records to Cosmos DB (upsert)"
    
    $syncCount = 0
    $failCount = 0
    
    foreach ($flight in $syllabusData) {
        try {
            # Create a unique ID for the document (combination of CAPID and Syllabus)
            $documentId = "$($flight.CAPID)-$($flight.Syllabus)"
            
            # Create the document object for upsert
            $documentObject = [PSCustomObject]@{
                id = $documentId
                CAPID = $flight.CAPID
                Syllabus = $flight.Syllabus
                FirstFlight = $flight.FirstFlight
                LastUpdated = Get-Date -Format o
                SyncSource = "OFlights"
            }
            
            # Upsert to Cosmos DB (creates new or updates existing)
            $result = Save-CosmosDbItem -Item $documentObject `
                -ConnectionString $cosmosConfig.ConnectionString `
                -Database $cosmosConfig.Database `
                -Container $cosmosConfig.Container
            
            if ($result) {
                $syncCount++
            } else {
                $failCount++
            }
        } catch {
            Write-Log "Failed to sync flight record for CAPID $($flight.CAPID), Syllabus $($flight.Syllabus). Error: $($_.Exception.Message)"
            $failCount++
            continue
        }
    }
    
    Write-Log "Cosmos DB sync completed. Successfully synced: $syncCount, Failed: $failCount"
    
} catch {
    Write-Log "Failed to connect to Cosmos DB. Error: $($_.Exception.Message)"
    exit 1
}

Write-Log "OFlights function completed successfully"
