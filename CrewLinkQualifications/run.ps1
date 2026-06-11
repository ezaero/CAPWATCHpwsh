<#
.SYNOPSIS
    Syncs CREWLINK-relevant CAPWATCH achievement qualifications to Cosmos DB.

.DESCRIPTION
    Reads Achievements.txt and MbrAchievements.txt from the downloaded CAPWATCH
    folder, joins member records to achievement names by AchvID, derives
    matching eligibility from status, and upserts qualification snapshots for
    CREWLINK.

.NOTES
    Runs on the same schedule as OFlights: 1400 UTC every Monday and Thursday.
    Output documents are partitioned by CAPID and use id format CAPID-AchvID.
#>

param($Timer)

$ErrorActionPreference = "Stop"
$CAPWATCHDATADIR = "$($env:HOME)\data\CAPWatch"
$locationPushed = $false

# Include shared Functions
. "$PSScriptRoot\..\shared\shared.ps1"
. "$PSScriptRoot\CrewLinkQualificationTransform.ps1"

function Get-CrewLinkQualificationsContainerName {
    if (-not [string]::IsNullOrWhiteSpace($env:CrewLinkQualificationsCosmosDbContainer)) {
        return $env:CrewLinkQualificationsCosmosDbContainer
    }

    return "crewlinkQualifications"
}

function Import-RequiredCsv {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (-not (Test-Path -Path $Path)) {
        throw "Required CAPWATCH file not found: $Name at $Path"
    }

    return Import-Csv $Path -ErrorAction Stop
}

function Query-CrewLinkCosmosDbContainer {
    param (
        [string]$ConnectionString,
        [string]$Database,
        [string]$Container,
        [string]$Query
    )

    try {
        $connStringParts = @{}
        $ConnectionString -split ';' | Where-Object { $_ -match '=' } | ForEach-Object {
            $key, $value = $_ -split '=', 2
            $connStringParts[$key.Trim()] = $value.Trim()
        }

        $endpoint = $connStringParts['AccountEndpoint'].TrimEnd('/')
        $key = $connStringParts['AccountKey']

        if (-not $endpoint -or -not $key) {
            Write-Log "Failed to parse Cosmos DB connection string for query"
            return @()
        }

        $uri = "$endpoint/dbs/$Database/colls/$Container/docs"
        $queryBody = @{ query = $Query } | ConvertTo-Json
        $allDocuments = @()
        $continuationToken = $null

        do {
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

            $responseHeaders = @{}
            $responseData = Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $queryBody -ContentType "application/query+json" -ResponseHeadersVariable responseHeaders -ErrorAction Stop

            if ($responseData.Documents) {
                $allDocuments += $responseData.Documents
            }

            $continuationToken = $null
            if ($responseHeaders.ContainsKey('x-ms-continuation')) {
                $continuationToken = $responseHeaders['x-ms-continuation']
                if ($continuationToken -is [array]) {
                    $continuationToken = $continuationToken[0]
                }
            }
        } while ($continuationToken)

        return $allDocuments
    } catch {
        Write-Log "Failed to query CrewLinkQualifications Cosmos DB container. Error: $($_.Exception.Message)"
        throw
    }
}

function Remove-CrewLinkCosmosDbDocument {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Document,

        [string]$ConnectionString,
        [string]$Database,
        [string]$Container
    )

    try {
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

        $documentId = $Document.id
        $partitionKeyValue = $Document.CAPID
        $uri = "$endpoint/dbs/$Database/colls/$Container/docs/$documentId"

        $verb = "delete"
        $resourceType = "docs"
        $resourceId = "dbs/$Database/colls/$Container/docs/$documentId"
        $date = [DateTime]::UtcNow.ToString('r')
        $stringToSign = "$verb`n$resourceType`n$resourceId`n$($date.ToLowerInvariant())`n`n"

        $hmacsha = New-Object System.Security.Cryptography.HMACSHA256
        $hmacsha.Key = [System.Convert]::FromBase64String($key)
        $hashBytes = $hmacsha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($stringToSign))
        $signature = [System.Convert]::ToBase64String($hashBytes)

        $authString = "type=master&ver=1.0&sig=$signature"
        $authToken = [System.Web.HttpUtility]::UrlEncode($authString)

        $headers = @{
            "Authorization"                = $authToken
            "x-ms-date"                    = $date
            "x-ms-version"                 = "2020-07-15"
            "x-ms-documentdb-partitionkey" = "[`"$partitionKeyValue`"]"
        }

        Invoke-RestMethod -Method DELETE -Uri $uri -Headers $headers -ErrorAction Stop | Out-Null
        return $true
    } catch {
        Write-Log "Failed to delete stale CrewLinkQualifications document $($Document.id). Error: $($_.Exception.Message)"
        return $false
    }
}

try {
    Push-Location $CAPWATCHDATADIR
    $locationPushed = $true

    $syncTime = Get-Date -Format o
    Write-Log "Starting CrewLinkQualifications function"

    $achievementsPath = Join-Path $CAPWATCHDATADIR "Achievements.txt"
    $memberAchievementsPath = Join-Path $CAPWATCHDATADIR "MbrAchievements.txt"

    $achievements = @(Import-RequiredCsv -Path $achievementsPath -Name "Achievements.txt")
    $memberAchievements = @(Import-RequiredCsv -Path $memberAchievementsPath -Name "MbrAchievements.txt")

    Write-Log "Imported $($achievements.Count) achievement lookup rows and $($memberAchievements.Count) member achievement rows"

    $qualificationSnapshots = @(ConvertTo-CrewLinkQualificationSnapshot `
        -MemberAchievements $memberAchievements `
        -Achievements $achievements `
        -SyncTime $syncTime)

    if ($qualificationSnapshots.Count -eq 0) {
        Write-Log "No CREWLINK qualification records found to sync. Exiting."
        return
    }

    $cosmosConfig = Get-CosmosDbConnection -Container (Get-CrewLinkQualificationsContainerName)
    if (-not $cosmosConfig.ConnectionString -or -not $cosmosConfig.Database -or -not $cosmosConfig.Container) {
        throw "Cosmos DB configuration incomplete. Verify CosmosDbConnectionString, CosmosDbDatabase, and CrewLinkQualificationsCosmosDbContainer."
    }

    $statusSummary = $qualificationSnapshots |
        Group-Object -Property EligibilityStatus |
        ForEach-Object { "$($_.Name): $($_.Count)" }
    Write-Log "Prepared $($qualificationSnapshots.Count) qualification snapshots. Eligibility summary: $($statusSummary -join ', ')"

    $existingSnapshots = @(Query-CrewLinkCosmosDbContainer `
        -ConnectionString $cosmosConfig.ConnectionString `
        -Database $cosmosConfig.Database `
        -Container $cosmosConfig.Container `
        -Query "SELECT * FROM c WHERE c.SyncSource = 'CrewLinkQualifications'")

    $syncPlan = Compare-CrewLinkQualificationSnapshots -DesiredSnapshots $qualificationSnapshots -ExistingSnapshots $existingSnapshots
    Write-Log "Sync plan: $($syncPlan.ToUpsert.Count) documents to upsert, $($syncPlan.ToDelete.Count) stale documents to delete, $($syncPlan.UnchangedCount) unchanged"

    $syncCount = 0
    $deleteCount = 0
    $failCount = 0

    foreach ($snapshot in $syncPlan.ToUpsert) {
        try {
            $snapshot.LastUpdated = Get-Date -Format o
            $result = Save-CosmosDbItem -Item $snapshot `
                -ConnectionString $cosmosConfig.ConnectionString `
                -Database $cosmosConfig.Database `
                -Container $cosmosConfig.Container

            if ($result) {
                $syncCount++
            } else {
                $failCount++
            }
        } catch {
            Write-Log "Failed to sync CREWLINK qualification for CAPID $($snapshot.CAPID), AchvID $($snapshot.AchvID). Error: $($_.Exception.Message)"
            $failCount++
        }
    }

    foreach ($staleSnapshot in $syncPlan.ToDelete) {
        $deleted = Remove-CrewLinkCosmosDbDocument -Document $staleSnapshot `
            -ConnectionString $cosmosConfig.ConnectionString `
            -Database $cosmosConfig.Database `
            -Container $cosmosConfig.Container

        if ($deleted) {
            $deleteCount++
        } else {
            $failCount++
        }
    }

    Write-Log "CrewLinkQualifications sync completed. Upserted: $syncCount, Deleted: $deleteCount, Unchanged: $($syncPlan.UnchangedCount), Failed: $failCount"

    if ($failCount -gt 0) {
        throw "CrewLinkQualifications sync completed with $failCount failed write/delete operations"
    }
} catch {
    Write-Log "CrewLinkQualifications function failed. Error: $($_.Exception.Message)"
    throw
} finally {
    if ($locationPushed) {
        Pop-Location
    }
}
