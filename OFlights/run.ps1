<#
.SYNOPSIS
    Processes O-Flight data from CAPWATCH and syncs it to Cosmos DB.

.DESCRIPTION
    This script performs the following tasks:
    1. Imports O-Flight records from the OFlight.txt file
    2. Filters flights for specific syllabuses (6, 7, 8, 9, 10)
    3. Groups flights by CAPID and Syllabus to identify first flight dates
    4. Validates that CAPIDs exist in the current Azure AD user list
    5. Deletes all existing O-Flight documents from Cosmos DB (source of truth model)
    6. Syncs flight data to Cosmos DB by upserting validated records
    7. Logs all actions for auditing purposes

.PARAMETER None
    This function uses environment variables for configuration

.NOTES
    - Requires $CAPWATCHDATADIR environment variable to be set
    - Requires Cosmos DB configuration in environment variables
    - Requires Microsoft Graph connection to be established
    - Uses delete-and-recreate pattern: OFlight.txt is the source of truth
#>

param($Timer)

# Set working directory to folder with all CAPWATCH CSV Text Files
$CAPWATCHDATADIR = "$($env:HOME)\data\CAPWatch"
Push-Location $CAPWATCHDATADIR

# Include shared Functions
. "$PSScriptRoot\..\shared\shared.ps1"

# Helper function to query Cosmos DB container
function Query-CosmosDbContainer {
    param (
        [string]$ConnectionString,
        [string]$Database,
        [string]$Container,
        [string]$Query
    )

    try {
        # Parse connection string
        $connStringParts = @{}
        $ConnectionString -split ';' | Where-Object { $_ -match '=' } | ForEach-Object {
            $key, $value = $_ -split '=', 2
            $connStringParts[$key.Trim()] = $value.Trim()
        }

        $endpoint = $connStringParts['AccountEndpoint'].TrimEnd('/')
        $key = $connStringParts['AccountKey']

        if (-not $endpoint -or -not $key) {
            Write-Log "Failed to parse Cosmos DB connection string"
            return @()
        }

        # Build URI for query
        $uri = "$endpoint/dbs/$Database/colls/$Container/docs"

        # Query body
        $queryBody = @{
            query = $Query
        } | ConvertTo-Json

        # Collect all documents across pages
        $allDocuments = @()
        $continuationToken = $null

        do {
            # Generate auth header for this request
            $verb = "post"
            $resourceType = "docs"
            $resourceId = "dbs/$Database/colls/$Container"
            $date = [DateTime]::UtcNow.ToString('r')

            $stringToSign = "$verb`n$resourceType`n$resourceId`n$($date.ToLowerInvariant())`n`n"

            $hmacsha = New-Object System.Security.Cryptography.HMACSHA256
            $hmacsha.Key = [System.Convert]::FromBase64String($key)
            $hashBytes = $hmacsha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($stringToSign))
            $signature = [System.Convert]::ToBase64String($hashBytes)

            $authString = "type=master&ver=1.0&sig=$signature"
            $authToken = [System.Web.HttpUtility]::UrlEncode($authString)

            # Build headers with continuation token if available
            $headers = @{
                "Authorization" = $authToken
                "x-ms-date" = $date
                "x-ms-version" = "2020-07-15"
                "x-ms-documentdb-isquery" = "true"
                "x-ms-documentdb-query-enablecrosspartition" = "true"
                "x-ms-max-item-count" = "1000"
            }

            if ($continuationToken) {
                $headers["x-ms-continuation"] = $continuationToken
            }

            # Execute query
            $responseHeaders = @{}
            $responseData = Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $queryBody -ContentType "application/query+json" -ResponseHeadersVariable responseHeaders -ErrorAction Stop

            # Add documents from this page
            if ($responseData.Documents) {
                $allDocuments += $responseData.Documents
            }

            # Get continuation token for next page
            $continuationToken = $null
            if ($responseHeaders.ContainsKey('x-ms-continuation')) {
                $continuationToken = $responseHeaders['x-ms-continuation']
                if ($continuationToken -is [array]) {
                    $continuationToken = $continuationToken[0]
                }
            }

        } while ($continuationToken)

        Write-Log "Queried $($allDocuments.Count) O-Flight documents from Cosmos DB"
        return $allDocuments

    } catch {
        Write-Log "Failed to query Cosmos DB container. Error: $($_.Exception.Message)"
        return @()
    }
}

# Helper function to delete a document from Cosmos DB
function Delete-CosmosDbDocument {
    param (
        [Parameter(Mandatory=$true)]
        [string]$DocumentId,
        [Parameter(Mandatory=$true)]
        [string]$PartitionKeyValue,
        [string]$ConnectionString,
        [string]$Database,
        [string]$Container
    )

    try {
        # Parse connection string
        $connStringParts = @{}
        $ConnectionString -split ';' | Where-Object { $_ -match '=' } | ForEach-Object {
            $key, $value = $_ -split '=', 2
            $connStringParts[$key.Trim()] = $value.Trim()
        }

        $endpoint = $connStringParts['AccountEndpoint'].TrimEnd('/')
        $key = $connStringParts['AccountKey']

        if (-not $endpoint -or -not $key) {
            Write-Log "Failed to parse Cosmos DB connection string for delete"
            return $false
        }

        # Build URI for delete
        $uri = "$endpoint/dbs/$Database/colls/$Container/docs/$DocumentId"

        # Generate auth header - Cosmos DB REST API requires specific format
        $verb = "delete"
        $resourceType = "docs"
        $resourceId = "dbs/$Database/colls/$Container/docs/$DocumentId"
        $date = [DateTime]::UtcNow.ToString('r')

        $stringToSign = "$verb`n$resourceType`n$resourceId`n$($date.ToLowerInvariant())`n`n"

        $hmacsha = New-Object System.Security.Cryptography.HMACSHA256
        $hmacsha.Key = [System.Convert]::FromBase64String($key)
        $hashBytes = $hmacsha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($stringToSign))
        $signature = [System.Convert]::ToBase64String($hashBytes)

        $authString = "type=master&ver=1.0&sig=$signature"
        $authToken = [System.Web.HttpUtility]::UrlEncode($authString)

        $headers = @{
            "Authorization"                  = $authToken
            "x-ms-date"                      = $date
            "x-ms-version"                   = "2020-07-15"
            "x-ms-documentdb-partitionkey"   = "[`"$PartitionKeyValue`"]"
        }

        Invoke-RestMethod -Method DELETE -Uri $uri -Headers $headers -ErrorAction Stop | Out-Null
        return $true

    } catch {
        Write-Log "Failed to delete document $DocumentId from Cosmos DB. Error: $($_.Exception.Message)"
        return $false
    }
}

# Helper function to batch delete documents in parallel
function Delete-DocumentsBatch {
    param (
        [Parameter(Mandatory=$true)]
        [array]$DocumentIds,
        [string]$ConnectionString,
        [string]$Database,
        [string]$Container,
        [int]$TimeoutSec = 30
    )

    if ($DocumentIds.Count -eq 0) {
        return 0
    }

    $deletedCount = 0
    
    # Delete in parallel (throttle to 5 concurrent)
    $DocumentIds | ForEach-Object -Parallel {
        $docId = $_
        $conn = $using:ConnectionString
        $db = $using:Database
        $cont = $using:Container
        $timeout = $using:TimeoutSec
        
        # Parse connection string in parallel job
        $connStringParts = @{}
        $conn -split ';' | Where-Object { $_ -match '=' } | ForEach-Object {
            $key, $value = $_ -split '=', 2
            $connStringParts[$key.Trim()] = $value.Trim()
        }

        $endpoint = $connStringParts['AccountEndpoint'].TrimEnd('/')
        $key = $connStringParts['AccountKey']
        
        # Extract CAPID from document ID (format: CAPID-Syllabus)
        $parts = $docId -split '-'
        $capId = $parts[0]
        
        $uri = "$endpoint/dbs/$db/colls/$cont/docs/$docId"

        try {
            # Generate auth header
            $verb = "delete"
            $resourceType = "docs"
            $resourceId = "dbs/$db/colls/$cont/docs/$docId"
            $date = [DateTime]::UtcNow.ToString('r')

            $stringToSign = "$verb`n$resourceType`n$resourceId`n$($date.ToLowerInvariant())`n`n"

            $hmacsha = New-Object System.Security.Cryptography.HMACSHA256
            $hmacsha.Key = [System.Convert]::FromBase64String($key)
            $hashBytes = $hmacsha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($stringToSign))
            $signature = [System.Convert]::ToBase64String($hashBytes)

            $authString = "type=master&ver=1.0&sig=$signature"
            $authToken = [System.Web.HttpUtility]::UrlEncode($authString)

            $headers = @{
                "Authorization"                  = $authToken
                "x-ms-date"                      = $date
                "x-ms-version"                   = "2020-07-15"
                "x-ms-documentdb-partitionkey"   = "[`"$capId`"]"
            }

            Invoke-RestMethod -Method DELETE -Uri $uri -Headers $headers -TimeoutSec $timeout -ErrorAction Stop | Out-Null
            [PSCustomObject]@{ Success = $true; DocId = $docId }
        } catch {
            [PSCustomObject]@{ Success = $false; DocId = $docId }
        }
    } -ThrottleLimit 5 | ForEach-Object {
        if ($_.Success) {
            $deletedCount++
        }
    }
    
    return $deletedCount
}

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

# Import the Training CSV file for Aircraft Ground Handling completion status
$allTrainingRecords = @()
$shouldSyncAircraftGroundHandling = $false
try {
    $trainingFilePath = "$($CAPWATCHDATADIR)\Training.txt"
    if (Test-Path -Path $trainingFilePath) {
        $allTrainingRecords = Import-Csv $trainingFilePath -ErrorAction Stop
        $shouldSyncAircraftGroundHandling = $true
        Write-Log "Successfully imported $($allTrainingRecords.Count) training records from Training.txt"
    } else {
        Write-Log "Training.txt not found at $trainingFilePath. Aircraft Ground Handling sync will be skipped."
    }
} catch {
    Write-Log "Failed to import Training.txt. Error: $($_.Exception.Message)"
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
# 2. Exclude future-dated flights; scheduled events are not completed O-Flights
# 3. Group by CAPID and Syllabus to get first flight
# 4. Filter out invalid CAPIDs
$today = (Get-Date).Date
$syllabusData = $allFlightRecords | 
    Where-Object { $_.Syllabus -in @("6", "7", "8", "9", "10") } |
    Where-Object {
        $flightDate = [datetime]::MinValue
        [datetime]::TryParse($_.FltDate, [ref]$flightDate) -and $flightDate.Date -le $today
    } |
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

$aircraftGroundHandlingLookup = @{}
$aircraftGroundHandlingData = @()
if ($shouldSyncAircraftGroundHandling) {
    $allTrainingRecords |
        Where-Object { $_.TypeCrs -and $_.TypeCrs.Trim() -eq "Aircraft Ground Handling" } |
        Where-Object { $_.CAPID -and $_.CAPID.Trim() -notmatch 'P$' } |
        Where-Object { $validEmployeeIds.ContainsKey($_.CAPID) } |
        Group-Object -Property CAPID |
        ForEach-Object {
            $firstCompletion = $_.Group | Sort-Object {
                $completedDate = [datetime]::MaxValue
                if ([datetime]::TryParse($_.Completed, [ref]$completedDate)) {
                    $completedDate
                } else {
                    [datetime]::MaxValue
                }
            } | Select-Object -First 1

            $aircraftGroundHandlingLookup[$_.Name] = $firstCompletion.Completed
        }

    $aircraftGroundHandlingData = foreach ($capid in $validEmployeeIds.Keys) {
        if ($capid -match 'P$') {
            continue
        }

        [PSCustomObject]@{
            CAPID = $capid
            Completed = if ($aircraftGroundHandlingLookup.ContainsKey($capid)) { $aircraftGroundHandlingLookup[$capid] } else { $null }
            Current = $aircraftGroundHandlingLookup.ContainsKey($capid)
        }
    }
}

if ($shouldSyncAircraftGroundHandling) {
    Write-Log "Processed training data: Found $($aircraftGroundHandlingLookup.Count) valid Aircraft Ground Handling completion records across $($aircraftGroundHandlingData.Count) known CAPIDs"
}

# Sync to Cosmos DB
if ($syllabusData.Count -eq 0 -and $aircraftGroundHandlingData.Count -eq 0) {
    Write-Log "No valid O-Flight or Aircraft Ground Handling records to sync. Exiting."
    exit 0
}

try {
    $cosmosConfig = Get-CosmosDbConnection
    
    if (-not $cosmosConfig.ConnectionString -or -not $cosmosConfig.Database -or -not $cosmosConfig.Container) {
        Write-Log "Error: Cosmos DB configuration incomplete. Verify CosmosDbConnectionString, CosmosDbDatabase, and CosmosDbContainer environment variables."
        exit 1
    }

    # Step 1: Query existing O-Flight documents from Cosmos DB
    Write-Log "Querying Cosmos DB for existing O-Flight documents..."
    $existingDocs = Query-CosmosDbContainer -ConnectionString $cosmosConfig.ConnectionString `
                                            -Database $cosmosConfig.Database `
                                            -Container $cosmosConfig.Container `
                                            -Query "SELECT c.id, c.CAPID, c.Syllabus, c.FirstFlight FROM c WHERE c.SyncSource = 'OFlights'"
    
    Write-Log "Found $($existingDocs.Count) existing O-Flight documents in Cosmos DB"
    
    # Step 2: Build lookup dictionaries for comparison
    # Create hash of existing documents keyed by ID
    $existingDocLookup = @{}
    foreach ($doc in $existingDocs) {
        $existingDocLookup[$doc.id] = $doc
    }
    
    # Create hash of expected documents keyed by ID
    $expectedDocLookup = @{}
    foreach ($flight in $syllabusData) {
        $docId = "$($flight.CAPID)-$($flight.Syllabus)"
        $expectedDocLookup[$docId] = $flight
    }
    
    Write-Log "Expected state: $($expectedDocLookup.Count) documents from OFlight.txt"
    
    # Step 3: Identify documents to delete (exist in Cosmos but not in source)
    $toDelete = @()
    foreach ($existingId in $existingDocLookup.Keys) {
        if (-not $expectedDocLookup.ContainsKey($existingId)) {
            $toDelete += $existingId
        }
    }
    
    # Step 4: Identify documents to upsert (new or changed)
    $toUpsert = @()
    foreach ($expectedId in $expectedDocLookup.Keys) {
        $expectedFlight = $expectedDocLookup[$expectedId]
        
        if ($existingDocLookup.ContainsKey($expectedId)) {
            # Document exists - check if FirstFlight changed
            $existingDoc = $existingDocLookup[$expectedId]
            if ($existingDoc.FirstFlight -ne $expectedFlight.FirstFlight) {
                $toUpsert += $expectedFlight
            }
            # else: Document unchanged, skip it
        } else {
            # Document is new
            $toUpsert += $expectedFlight
        }
    }
    
    Write-Log "Sync plan: $($toDelete.Count) documents to delete, $($toUpsert.Count) documents to upsert, $($existingDocs.Count - $toDelete.Count - $toUpsert.Count) documents unchanged"
    
    # Step 5: Delete orphaned documents in parallel
    $deleteCount = 0
    if ($toDelete.Count -gt 0) {
        Write-Log "Deleting $($toDelete.Count) orphaned documents (removed from source)..."
        $deleteCount = Delete-DocumentsBatch -DocumentIds $toDelete `
                                             -ConnectionString $cosmosConfig.ConnectionString `
                                             -Database $cosmosConfig.Database `
                                             -Container $cosmosConfig.Container
        Write-Log "Successfully deleted $deleteCount orphaned documents. Failed: $($toDelete.Count - $deleteCount)"
    }
    
    # Step 6: Upsert new and changed documents
    Write-Log "Syncing $($toUpsert.Count) new and changed flight records to Cosmos DB"
    
    $syncCount = 0
    $failCount = 0
    
    foreach ($flight in $toUpsert) {
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
            
            # Upsert to Cosmos DB (creates new or updates changed)
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
    
    Write-Log "Cosmos DB sync completed. Deleted: $deleteCount, Upserted: $syncCount, Failed: $failCount. Total unchanged (not synced): $($existingDocs.Count - $toDelete.Count - $toUpsert.Count)"

    # Step 7: Upsert Aircraft Ground Handling status records. This avoids deleting
    # records while still making absent Training.txt rows show as not current.
    if (-not $shouldSyncAircraftGroundHandling) {
        Write-Log "Aircraft Ground Handling sync skipped because Training.txt was not available"
    } else {
        Write-Log "Querying Cosmos DB for existing Aircraft Ground Handling documents..."
        $existingAghDocs = Query-CosmosDbContainer -ConnectionString $cosmosConfig.ConnectionString `
                                                   -Database $cosmosConfig.Database `
                                                   -Container $cosmosConfig.Container `
                                                   -Query "SELECT c.id, c.CAPID, c.Completed, c.Current FROM c WHERE c.SyncSource = 'AircraftGroundHandling'"

        $parentAghDocumentIds = @(
            $existingAghDocs |
                Where-Object { $_.CAPID -and $_.CAPID.Trim() -match 'P$' } |
                ForEach-Object { $_.id }
        )

        $existingAghDocLookup = @{}
        foreach ($doc in $existingAghDocs) {
            if ($doc.CAPID -and $doc.CAPID.Trim() -notmatch 'P$') {
                $existingAghDocLookup[$doc.id] = $doc
            }
        }

        $aghToUpsert = @()
        foreach ($training in $aircraftGroundHandlingData) {
            $documentId = "$($training.CAPID)-AircraftGroundHandling"

            if ($existingAghDocLookup.ContainsKey($documentId)) {
                $existingDoc = $existingAghDocLookup[$documentId]
                if ($existingDoc.Completed -ne $training.Completed -or [bool]$existingDoc.Current -ne [bool]$training.Current) {
                    $aghToUpsert += $training
                }
            } else {
                $aghToUpsert += $training
            }
        }

        Write-Log "Aircraft Ground Handling sync plan: $($parentAghDocumentIds.Count) parent documents to delete, $($aghToUpsert.Count) status records to upsert, $($aircraftGroundHandlingData.Count - $aghToUpsert.Count) records unchanged"

        $aghParentDeleteCount = 0
        if ($parentAghDocumentIds.Count -gt 0) {
            Write-Log "Deleting $($parentAghDocumentIds.Count) parent Aircraft Ground Handling documents..."
            $aghParentDeleteCount = Delete-DocumentsBatch -DocumentIds $parentAghDocumentIds `
                                                          -ConnectionString $cosmosConfig.ConnectionString `
                                                          -Database $cosmosConfig.Database `
                                                          -Container $cosmosConfig.Container
        }

        $aghSyncCount = 0
        $aghFailCount = 0

        Write-Log "Syncing $($aghToUpsert.Count) new and changed Aircraft Ground Handling status records to Cosmos DB"

        foreach ($training in $aghToUpsert) {
            try {
                $documentId = "$($training.CAPID)-AircraftGroundHandling"
                $documentObject = [PSCustomObject]@{
                    id = $documentId
                    CAPID = $training.CAPID
                    TypeCrs = "Aircraft Ground Handling"
                    Completed = $training.Completed
                    Current = $training.Current
                    LastUpdated = Get-Date -Format o
                    SyncSource = "AircraftGroundHandling"
                }

                $result = Save-CosmosDbItem -Item $documentObject `
                    -ConnectionString $cosmosConfig.ConnectionString `
                    -Database $cosmosConfig.Database `
                    -Container $cosmosConfig.Container

                if ($result) {
                    $aghSyncCount++
                } else {
                    $aghFailCount++
                }
            } catch {
                Write-Log "Failed to sync Aircraft Ground Handling record for CAPID $($training.CAPID). Error: $($_.Exception.Message)"
                $aghFailCount++
                continue
            }
        }

        Write-Log "Aircraft Ground Handling sync completed. Deleted parent docs: $aghParentDeleteCount, Upserted: $aghSyncCount, Failed: $aghFailCount"
    }
    
} catch {
    Write-Log "Failed to connect to Cosmos DB. Error: $($_.Exception.Message)"
    exit 1
}

Write-Log "OFlights function completed successfully"
