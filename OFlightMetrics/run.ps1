<#
.SYNOPSIS
    Calculates Orientation Flight metrics and priority scores by squadron.

.DESCRIPTION
    This script performs the following tasks:
    1. Runs every Monday and Thursday at 8:00 AM MST
    2. Queries Cosmos DB syllabus container for all O-Flight records
    3. Maps CAPIDs to squadrons using Azure AD user data
    4. Calculates metrics for:
       - Previous month (e.g., if run on Feb 1, calculates Jan metrics)
       - Current Fiscal Year (since Oct 1)
    5. Writes metrics to Cosmos DB metrics container
    6. Logs all actions for auditing

.NOTES
    Requires Cosmos DB and Microsoft Graph API configuration in environment variables
    Fiscal Year runs from October 1 to September 30
#>

param($Timer)

# Include shared functions
. "$PSScriptRoot\..\shared\shared.ps1"

$logPrefix = '📊 [OFlightMetrics]'

Write-Log "$logPrefix Function started at $(Get-Date -Format o)"

# Helper function to query Cosmos DB container with pagination support
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
            Write-Log "$logPrefix Failed to parse Cosmos DB connection string"
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
        $pageCount = 0

        do {
            $pageCount++

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

            # Execute query using Invoke-RestMethod with response headers
            try {
                $responseHeaders = @{}
                $responseData = Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $queryBody -ContentType "application/query+json" -ResponseHeadersVariable responseHeaders -ErrorAction Stop

                # Add documents from this page
                if ($responseData.Documents) {
                    $allDocuments += $responseData.Documents
                    Write-Log "$logPrefix   Retrieved page $pageCount : $($responseData.Documents.Count) documents (total so far: $($allDocuments.Count))"
                }

                # Get continuation token for next page
                $continuationToken = $null
                if ($responseHeaders.ContainsKey('x-ms-continuation')) {
                    $continuationToken = $responseHeaders['x-ms-continuation']
                    if ($continuationToken -is [array]) {
                        $continuationToken = $continuationToken[0]
                    }
                }

                # If no documents returned, we're done
                if (-not $responseData.Documents -or $responseData.Documents.Count -eq 0) {
                    break
                }

            } catch {
                $errorDetails = ""
                if ($_.ErrorDetails) {
                    $errorDetails = $_.ErrorDetails.Message
                }
                Write-Log "$logPrefix   Error on page $pageCount : $($_.Exception.Message) $errorDetails"

                # If this is page 1, rethrow the error
                if ($pageCount -eq 1) {
                    throw
                }

                # Otherwise, stop pagination and return what we have
                Write-Log "$logPrefix   Stopping pagination due to error. Returning $($allDocuments.Count) documents collected so far."
                break
            }

        } while ($continuationToken)

        Write-Log "$logPrefix   Query complete: $($allDocuments.Count) total documents retrieved across $pageCount page(s)"
        return $allDocuments

    } catch {
        Write-Log "$logPrefix Failed to query Cosmos DB container $Container. Error: $($_.Exception.Message)"
        return @()
    }
}

# Helper function to save metrics to Cosmos DB metrics container
function Save-MetricsItem {
    param (
        [Parameter(Mandatory=$true)]
        [object]$Item,
        [string]$ConnectionString = $env:CosmosDbConnectionString,
        [string]$Database = $env:CosmosDbDatabase
    )

    try {
        # Ensure item has required id field
        if (-not $Item.id) {
            Write-Log "$logPrefix Error: Metrics item must have an 'id' field"
            return $false
        }

        # Ensure item has metricType for partition key
        if (-not $Item.metricType) {
            Write-Log "$logPrefix Error: Metrics item must have a 'metricType' field for partition key"
            return $false
        }

        # Add timestamp if not present
        if (-not $Item.calculatedAt) {
            $Item | Add-Member -MemberType NoteProperty -Name "calculatedAt" -Value (Get-Date -Format o) -ErrorAction SilentlyContinue
        }

        # Parse connection string
        $connStringParts = @{}
        $ConnectionString -split ';' | Where-Object { $_ -match '=' } | ForEach-Object {
            $key, $value = $_ -split '=', 2
            $connStringParts[$key.Trim()] = $value.Trim()
        }

        $endpoint = $connStringParts['AccountEndpoint'].TrimEnd('/')
        $key = $connStringParts['AccountKey']

        if (-not $endpoint -or -not $key) {
            Write-Log "$logPrefix Failed to parse Cosmos DB connection string"
            return $false
        }

        # Build URI for upsert (metrics container)
        $uri = "$endpoint/dbs/$Database/colls/metrics/docs"

        # Generate auth header
        $verb = "post"
        $resourceType = "docs"
        $resourceId = "dbs/$Database/colls/metrics"
        $date = [DateTime]::UtcNow.ToString('r')

        $stringToSign = "$verb`n$resourceType`n$resourceId`n$($date.ToLowerInvariant())`n`n"

        $hmacsha = New-Object System.Security.Cryptography.HMACSHA256
        $hmacsha.Key = [System.Convert]::FromBase64String($key)
        $hashBytes = $hmacsha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($stringToSign))
        $signature = [System.Convert]::ToBase64String($hashBytes)

        $authString = "type=master&ver=1.0&sig=$signature"
        $authToken = [System.Web.HttpUtility]::UrlEncode($authString)

        # Use metricType as partition key
        $partitionKeyValue = $Item.metricType

        $headers = @{
            "Authorization"                  = $authToken
            "x-ms-date"                      = $date
            "x-ms-version"                   = "2020-07-15"
            "x-ms-documentdb-is-upsert"      = "true"
            "x-ms-documentdb-partitionkey"   = "[`"$partitionKeyValue`"]"
        }

        $body = $Item | ConvertTo-Json -Depth 10

        $response = Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $body -ContentType "application/json" -ErrorAction Stop

        Write-Log "$logPrefix Successfully saved metrics to Cosmos DB: $($Item.id)"
        return $true

    } catch {
        $errorMessage = $_.Exception.Message
        Write-Log "$logPrefix Failed to save metrics to Cosmos DB: $($Item.id). Error: $errorMessage"
        return $false
    }
}

try {
    # Authenticate to Microsoft Graph using managed identity
    try {
        $MSGraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token
        Connect-MgGraph -AccessToken $MSGraphAccessToken -NoWelcome
        Write-Log "$logPrefix Successfully authenticated to Microsoft Graph"
    } catch {
        Write-Log "$logPrefix Failed to authenticate to Microsoft Graph. Error: $($_.Exception.Message)"
        throw
    }

    # Get Cosmos DB configuration
    $cosmosConfig = Get-CosmosDbConnection
    $cosmosConnectionString = $cosmosConfig.ConnectionString
    $cosmosDatabase = $cosmosConfig.Database

    if (-not $cosmosConnectionString -or -not $cosmosDatabase) {
        Write-Log "$logPrefix Error: Cosmos DB configuration incomplete"
        throw "Cosmos DB configuration incomplete"
    }

    # Calculate date ranges
    $now = Get-Date

    # Previous month: First day of last month to last day of last month
    $previousMonthStart = (Get-Date -Day 1).AddMonths(-1)
    $previousMonthEnd = (Get-Date -Day 1).AddDays(-1)
    $previousMonthName = $previousMonthStart.ToString('MMMM yyyy')

    # Fiscal Year: October 1 of the fiscal year to today
    $currentYear = $now.Year
    $currentMonth = $now.Month
    if ($currentMonth -ge 10) {
        # Oct-Dec: FY starts Oct 1 of current year
        $fyStart = Get-Date -Year $currentYear -Month 10 -Day 1 -Hour 0 -Minute 0 -Second 0
    } else {
        # Jan-Sep: FY starts Oct 1 of previous year
        $fyStart = Get-Date -Year ($currentYear - 1) -Month 10 -Day 1 -Hour 0 -Minute 0 -Second 0
    }
    $fyName = "FY$($fyStart.Year + 1)"

    Write-Log "$logPrefix Calculating metrics for:"
    Write-Log "$logPrefix   - Previous Month: $previousMonthName ($($previousMonthStart.ToString('yyyy-MM-dd')) to $($previousMonthEnd.ToString('yyyy-MM-dd')))"
    Write-Log "$logPrefix   - Fiscal Year: $fyName (since $($fyStart.ToString('yyyy-MM-dd')))"

    # Query all O-Flight records from syllabus container
    Write-Log "$logPrefix Querying syllabus container for all O-Flight records..."
    $allFlights = Query-CosmosDbContainer -ConnectionString $cosmosConnectionString `
                                          -Database $cosmosDatabase `
                                          -Container "syllabus" `
                                          -Query "SELECT * FROM c WHERE c.SyncSource = 'OFlights'"

    Write-Log "$logPrefix Retrieved $($allFlights.Count) O-Flight records from Cosmos DB"

    # Get all users to map CAPID to squadron
    Write-Log "$logPrefix Retrieving users from Azure AD..."
    $allUsers = GetAllUsers -SelectFields "mail,displayName,companyName,employeeId,id"
    Write-Log "$logPrefix Retrieved $($allUsers.Count) users from Azure AD"

    # Build CAPID to squadron mapping
    $capidToSquadron = @{}
    foreach ($user in $allUsers) {
        if ($user.employeeId -and $user.companyName) {
            # Extract squadron from companyName (format: "CO-XXX")
            if ($user.companyName -match 'CO-(.+)') {
                $squadron = $matches[1]
                $capidToSquadron[$user.employeeId] = $squadron
            }
        }
    }
    Write-Log "$logPrefix Mapped $($capidToSquadron.Count) CAPIDs to squadrons"

    # Filter and count flights for previous month
    $previousMonthFlights = $allFlights | Where-Object {
        try {
            $flightDate = [DateTime]::Parse($_.FirstFlight)
            $flightDate -ge $previousMonthStart -and $flightDate -le $previousMonthEnd
        } catch {
            $false
        }
    }

    Write-Log "$logPrefix Found $($previousMonthFlights.Count) flights in previous month"

    # Filter and count flights for fiscal year
    $fyFlights = $allFlights | Where-Object {
        try {
            $flightDate = [DateTime]::Parse($_.FirstFlight)
            $flightDate -ge $fyStart
        } catch {
            $false
        }
    }

    Write-Log "$logPrefix Found $($fyFlights.Count) flights in current fiscal year"

    # Group previous month flights by squadron
    $previousMonthBySquadron = @{}
    foreach ($flight in $previousMonthFlights) {
        $squadron = $capidToSquadron[$flight.CAPID]
        if ($squadron) {
            if (-not $previousMonthBySquadron.ContainsKey($squadron)) {
                $previousMonthBySquadron[$squadron] = @{
                    Total = 0
                    BySyllabus = @{
                        "6" = 0
                        "7" = 0
                        "8" = 0
                        "9" = 0
                        "10" = 0
                    }
                    UniqueCadets = @()
                }
            }
            $previousMonthBySquadron[$squadron].Total++
            $previousMonthBySquadron[$squadron].BySyllabus[$flight.Syllabus]++
            if ($previousMonthBySquadron[$squadron].UniqueCadets -notcontains $flight.CAPID) {
                $previousMonthBySquadron[$squadron].UniqueCadets += $flight.CAPID
            }
        }
    }

    # Group fiscal year flights by squadron
    $fyBySquadron = @{}
    foreach ($flight in $fyFlights) {
        $squadron = $capidToSquadron[$flight.CAPID]
        if ($squadron) {
            if (-not $fyBySquadron.ContainsKey($squadron)) {
                $fyBySquadron[$squadron] = @{
                    Total = 0
                    BySyllabus = @{
                        "6" = 0
                        "7" = 0
                        "8" = 0
                        "9" = 0
                        "10" = 0
                    }
                    UniqueCadets = @()
                }
            }
            $fyBySquadron[$squadron].Total++
            $fyBySquadron[$squadron].BySyllabus[$flight.Syllabus]++
            if ($fyBySquadron[$squadron].UniqueCadets -notcontains $flight.CAPID) {
                $fyBySquadron[$squadron].UniqueCadets += $flight.CAPID
            }
        }
    }

    # Save previous month metrics to Cosmos DB
    $previousMonthId = "monthly-$($previousMonthStart.ToString('yyyy-MM'))"
    $previousMonthMetric = [PSCustomObject]@{
        id = $previousMonthId
        metricType = "monthly-squadron"
        period = "month"
        year = $previousMonthStart.Year
        month = $previousMonthStart.Month
        monthName = $previousMonthName
        startDate = $previousMonthStart.ToString('yyyy-MM-dd')
        endDate = $previousMonthEnd.ToString('yyyy-MM-dd')
        totalFlights = $previousMonthFlights.Count
        totalUniqueCadets = ($previousMonthFlights | Select-Object -ExpandProperty CAPID -Unique).Count
        squadrons = @{}
        calculatedAt = (Get-Date -Format o)
    }

    # Add squadron details to previous month metric
    foreach ($squadron in $previousMonthBySquadron.Keys) {
        $previousMonthMetric.squadrons[$squadron] = @{
            totalFlights = $previousMonthBySquadron[$squadron].Total
            uniqueCadets = $previousMonthBySquadron[$squadron].UniqueCadets.Count
            bySyllabus = $previousMonthBySquadron[$squadron].BySyllabus
        }
    }

    Write-Log "$logPrefix Saving previous month metrics ($previousMonthName)..."
    $saved = Save-MetricsItem -Item $previousMonthMetric -ConnectionString $cosmosConnectionString -Database $cosmosDatabase
    if ($saved) {
        Write-Log "$logPrefix ✅ Previous month metrics saved successfully"
    } else {
        Write-Log "$logPrefix ❌ Failed to save previous month metrics"
    }

    # Save fiscal year metrics to Cosmos DB
    $fyId = "fiscal-year-$($fyStart.Year + 1)"
    $fyMetric = [PSCustomObject]@{
        id = $fyId
        metricType = "fiscal-year-squadron"
        period = "fiscal-year"
        fiscalYear = $fyStart.Year + 1
        fiscalYearName = $fyName
        startDate = $fyStart.ToString('yyyy-MM-dd')
        endDate = $now.ToString('yyyy-MM-dd')
        totalFlights = $fyFlights.Count
        totalUniqueCadets = ($fyFlights | Select-Object -ExpandProperty CAPID -Unique).Count
        squadrons = @{}
        calculatedAt = (Get-Date -Format o)
    }

    # Add squadron details to fiscal year metric
    foreach ($squadron in $fyBySquadron.Keys) {
        $fyMetric.squadrons[$squadron] = @{
            totalFlights = $fyBySquadron[$squadron].Total
            uniqueCadets = $fyBySquadron[$squadron].UniqueCadets.Count
            bySyllabus = $fyBySquadron[$squadron].BySyllabus
        }
    }

    Write-Log "$logPrefix Saving fiscal year metrics ($fyName)..."
    $saved = Save-MetricsItem -Item $fyMetric -ConnectionString $cosmosConnectionString -Database $cosmosDatabase
    if ($saved) {
        Write-Log "$logPrefix ✅ Fiscal year metrics saved successfully"
    } else {
        Write-Log "$logPrefix ❌ Failed to save fiscal year metrics"
    }

    # ============================================================================
    # ADDITIONAL METRICS CALCULATION
    # ============================================================================
    Write-Log "$logPrefix Calculating additional metrics..."

    # METRIC 1: Zero O-Flights per Squadron (Number & Percentage)
    Write-Log "$logPrefix   Calculating zero O-Flights metrics..."
    $zeroFlightsMetric = [PSCustomObject]@{
        id = "zero-flights-$($now.ToString('yyyy-MM'))"
        metricType = "zero-flights-squadron"
        calculatedDate = $now.ToString('yyyy-MM-dd')
        totalCadets = 0
        cadetsWithZeroFlights = 0
        percentageWithZeroFlights = 0
        squadrons = @{}
        calculatedAt = (Get-Date -Format o)
    }

    # Get all active cadets (employeeType = 'Cadet') grouped by squadron
    $activeCadetsBySquadron = @{}
    $cadetsWithFlights = $allFlights | Select-Object -ExpandProperty CAPID -Unique

    foreach ($user in $allUsers) {
        if ($user.employeeType -eq 'Cadet' -and $user.employeeId -and $user.companyName) {
            if ($user.companyName -match 'CO-(.+)') {
                $squadron = $matches[1]
                if (-not $activeCadetsBySquadron.ContainsKey($squadron)) {
                    $activeCadetsBySquadron[$squadron] = @{
                        TotalCadets = 0
                        CadetsWithZeroFlights = 0
                        CadetList = @()
                    }
                }
                $activeCadetsBySquadron[$squadron].TotalCadets++
                $activeCadetsBySquadron[$squadron].CadetList += $user.employeeId

                if ($cadetsWithFlights -notcontains $user.employeeId) {
                    $activeCadetsBySquadron[$squadron].CadetsWithZeroFlights++
                }
            }
        }
    }

    # Calculate totals and percentages
    $totalCadets = 0
    $totalWithZeroFlights = 0
    foreach ($squadron in $activeCadetsBySquadron.Keys) {
        $sqData = $activeCadetsBySquadron[$squadron]
        $totalCadets += $sqData.TotalCadets
        $totalWithZeroFlights += $sqData.CadetsWithZeroFlights

        $percentage = if ($sqData.TotalCadets -gt 0) {
            [math]::Round(($sqData.CadetsWithZeroFlights / $sqData.TotalCadets) * 100, 1)
        } else { 0 }

        $zeroFlightsMetric.squadrons[$squadron] = @{
            totalCadets = $sqData.TotalCadets
            cadetsWithZeroFlights = $sqData.CadetsWithZeroFlights
            percentageWithZeroFlights = $percentage
        }
    }

    $zeroFlightsMetric.totalCadets = $totalCadets
    $zeroFlightsMetric.cadetsWithZeroFlights = $totalWithZeroFlights
    $zeroFlightsMetric.percentageWithZeroFlights = if ($totalCadets -gt 0) {
        [math]::Round(($totalWithZeroFlights / $totalCadets) * 100, 1)
    } else { 0 }

    $saved = Save-MetricsItem -Item $zeroFlightsMetric -ConnectionString $cosmosConnectionString -Database $cosmosDatabase
    Write-Log "$logPrefix   ✅ Zero O-Flights metrics saved ($totalWithZeroFlights of $totalCadets cadets have zero flights)"

    # METRIC 2: Syllabus Completion Rate per Squadron
    Write-Log "$logPrefix   Calculating syllabus completion rates..."
    $syllabusCompletionMetric = [PSCustomObject]@{
        id = "syllabus-completion-$($now.ToString('yyyy-MM'))"
        metricType = "syllabus-completion"
        calculatedDate = $now.ToString('yyyy-MM-dd')
        wingWide = @{
            "6" = @{ completed = 0; percentage = 0 }
            "7" = @{ completed = 0; percentage = 0 }
            "8" = @{ completed = 0; percentage = 0 }
            "9" = @{ completed = 0; percentage = 0 }
            "10" = @{ completed = 0; percentage = 0 }
            "5for5" = @{ completed = 0; percentage = 0 }
        }
        squadrons = @{}
        calculatedAt = (Get-Date -Format o)
    }

    # Count cadets who completed each syllabus
    $cadetsBySyllabus = @{
        "6" = @()
        "7" = @()
        "8" = @()
        "9" = @()
        "10" = @()
    }

    foreach ($flight in $allFlights) {
        $syllabus = $flight.Syllabus
        if ($cadetsBySyllabus[$syllabus] -notcontains $flight.CAPID) {
            $cadetsBySyllabus[$syllabus] += $flight.CAPID
        }
    }

    # Wing-wide percentages
    $totalActiveCadets = ($allUsers | Where-Object { $_.employeeType -eq 'Cadet' -and $_.employeeId }).Count
    foreach ($syllabus in @("6", "7", "8", "9", "10")) {
        $completed = $cadetsBySyllabus[$syllabus].Count
        $syllabusCompletionMetric.wingWide[$syllabus].completed = $completed
        $syllabusCompletionMetric.wingWide[$syllabus].percentage = if ($totalActiveCadets -gt 0) {
            [math]::Round(($completed / $totalActiveCadets) * 100, 1)
        } else { 0 }
    }

    # 5-for-5 completers (all syllabuses)
    $fiveForFiveCadets = $cadetsBySyllabus["6"] | Where-Object {
        $cadetsBySyllabus["7"] -contains $_ -and
        $cadetsBySyllabus["8"] -contains $_ -and
        $cadetsBySyllabus["9"] -contains $_ -and
        $cadetsBySyllabus["10"] -contains $_
    }
    $syllabusCompletionMetric.wingWide["5for5"].completed = $fiveForFiveCadets.Count
    $syllabusCompletionMetric.wingWide["5for5"].percentage = if ($totalActiveCadets -gt 0) {
        [math]::Round(($fiveForFiveCadets.Count / $totalActiveCadets) * 100, 1)
    } else { 0 }

    # Per-squadron completion rates
    foreach ($squadron in $activeCadetsBySquadron.Keys) {
        $sqCadets = $activeCadetsBySquadron[$squadron].CadetList
        $sqTotal = $sqCadets.Count

        $syllabusCompletionMetric.squadrons[$squadron] = @{
            totalCadets = $sqTotal
            "6" = @{ completed = 0; percentage = 0 }
            "7" = @{ completed = 0; percentage = 0 }
            "8" = @{ completed = 0; percentage = 0 }
            "9" = @{ completed = 0; percentage = 0 }
            "10" = @{ completed = 0; percentage = 0 }
            "5for5" = @{ completed = 0; percentage = 0 }
        }

        foreach ($syllabus in @("6", "7", "8", "9", "10")) {
            $completed = ($sqCadets | Where-Object { $cadetsBySyllabus[$syllabus] -contains $_ }).Count
            $syllabusCompletionMetric.squadrons[$squadron][$syllabus].completed = $completed
            $syllabusCompletionMetric.squadrons[$squadron][$syllabus].percentage = if ($sqTotal -gt 0) {
                [math]::Round(($completed / $sqTotal) * 100, 1)
            } else { 0 }
        }

        # Squadron 5-for-5
        $sq5for5 = ($sqCadets | Where-Object { $fiveForFiveCadets -contains $_ }).Count
        $syllabusCompletionMetric.squadrons[$squadron]["5for5"].completed = $sq5for5
        $syllabusCompletionMetric.squadrons[$squadron]["5for5"].percentage = if ($sqTotal -gt 0) {
            [math]::Round(($sq5for5 / $sqTotal) * 100, 1)
        } else { 0 }
    }

    $saved = Save-MetricsItem -Item $syllabusCompletionMetric -ConnectionString $cosmosConnectionString -Database $cosmosDatabase
    Write-Log "$logPrefix   ✅ Syllabus completion metrics saved (5-for-5: $($fiveForFiveCadets.Count) cadets)"

    # METRIC 3: Monthly Flight Trends (Past 24 Months)
    Write-Log "$logPrefix   Calculating monthly flight trends (24 months)..."
    $monthlyTrendsMetric = [PSCustomObject]@{
        id = "monthly-trends-24"
        metricType = "monthly-trends"
        periodMonths = 24
        calculatedDate = $now.ToString('yyyy-MM-dd')
        months = @()
        calculatedAt = (Get-Date -Format o)
    }

    # Get flights from past 24 months
    $twoYearsAgo = $now.AddMonths(-24)
    for ($i = 0; $i -lt 24; $i++) {
        $monthStart = $twoYearsAgo.AddMonths($i)
        $monthEnd = $monthStart.AddMonths(1).AddDays(-1)

        $monthFlights = $allFlights | Where-Object {
            try {
                $flightDate = [DateTime]::Parse($_.FirstFlight)
                $flightDate -ge $monthStart -and $flightDate -le $monthEnd
            } catch {
                $false
            }
        }

        $monthlyTrendsMetric.months += [PSCustomObject]@{
            year = $monthStart.Year
            month = $monthStart.Month
            monthName = $monthStart.ToString('MMM yyyy')
            totalFlights = $monthFlights.Count
            uniqueCadets = ($monthFlights | Select-Object -ExpandProperty CAPID -Unique).Count
        }
    }

    $saved = Save-MetricsItem -Item $monthlyTrendsMetric -ConnectionString $cosmosConnectionString -Database $cosmosDatabase
    Write-Log "$logPrefix   ✅ Monthly trends saved (24 months of data)"

    # METRIC 4: Year-over-Year Comparison
    Write-Log "$logPrefix   Calculating year-over-year comparison..."

    # Previous FY same period
    $prevFyStart = $fyStart.AddYears(-1)
    $prevFyEnd = $now.AddYears(-1)

    $prevFyFlights = $allFlights | Where-Object {
        try {
            $flightDate = [DateTime]::Parse($_.FirstFlight)
            $flightDate -ge $prevFyStart -and $flightDate -le $prevFyEnd
        } catch {
            $false
        }
    }

    $yoyMetric = [PSCustomObject]@{
        id = "year-over-year-$($fyStart.Year + 1)"
        metricType = "year-over-year"
        currentFY = $fyStart.Year + 1
        previousFY = $fyStart.Year
        calculatedDate = $now.ToString('yyyy-MM-dd')
        current = @{
            fiscalYear = $fyStart.Year + 1
            totalFlights = $fyFlights.Count
            uniqueCadets = ($fyFlights | Select-Object -ExpandProperty CAPID -Unique).Count
            startDate = $fyStart.ToString('yyyy-MM-dd')
            endDate = $now.ToString('yyyy-MM-dd')
        }
        previous = @{
            fiscalYear = $fyStart.Year
            totalFlights = $prevFyFlights.Count
            uniqueCadets = ($prevFyFlights | Select-Object -ExpandProperty CAPID -Unique).Count
            startDate = $prevFyStart.ToString('yyyy-MM-dd')
            endDate = $prevFyEnd.ToString('yyyy-MM-dd')
        }
        change = @{
            flights = $fyFlights.Count - $prevFyFlights.Count
            flightsPercentage = if ($prevFyFlights.Count -gt 0) {
                [math]::Round((($fyFlights.Count - $prevFyFlights.Count) / $prevFyFlights.Count) * 100, 1)
            } else { 0 }
            cadets = ($fyFlights | Select-Object -ExpandProperty CAPID -Unique).Count - ($prevFyFlights | Select-Object -ExpandProperty CAPID -Unique).Count
        }
        calculatedAt = (Get-Date -Format o)
    }

    $saved = Save-MetricsItem -Item $yoyMetric -ConnectionString $cosmosConnectionString -Database $cosmosDatabase
    Write-Log "$logPrefix   ✅ Year-over-year comparison saved ($($yoyMetric.change.flightsPercentage)% change)"

    # METRIC 5: Squadron Rankings by Engagement Rate
    Write-Log "$logPrefix   Calculating squadron rankings..."
    $squadronRankings = @()

    foreach ($squadron in $activeCadetsBySquadron.Keys) {
        $sqData = $activeCadetsBySquadron[$squadron]
        $sqFlights = if ($fyBySquadron.ContainsKey($squadron)) { $fyBySquadron[$squadron].Total } else { 0 }
        $sqCadetsWithFlights = if ($fyBySquadron.ContainsKey($squadron)) { $fyBySquadron[$squadron].UniqueCadets.Count } else { 0 }

        $engagementRate = if ($sqData.TotalCadets -gt 0) {
            [math]::Round(($sqFlights / $sqData.TotalCadets), 2)
        } else { 0 }

        $participationRate = if ($sqData.TotalCadets -gt 0) {
            [math]::Round(($sqCadetsWithFlights / $sqData.TotalCadets) * 100, 1)
        } else { 0 }

        $squadronRankings += [PSCustomObject]@{
            squadron = $squadron
            totalCadets = $sqData.TotalCadets
            totalFlights = $sqFlights
            cadetsWithFlights = $sqCadetsWithFlights
            engagementRate = $engagementRate
            participationRate = $participationRate
        }
    }

    $squadronRankingsMetric = [PSCustomObject]@{
        id = "squadron-rankings-$($now.ToString('yyyy-MM'))"
        metricType = "squadron-rankings"
        fiscalYear = $fyStart.Year + 1
        calculatedDate = $now.ToString('yyyy-MM-dd')
        rankings = @{
            byEngagementRate = ($squadronRankings | Sort-Object -Property engagementRate -Descending | Select-Object -First 10)
            byParticipationRate = ($squadronRankings | Sort-Object -Property participationRate -Descending | Select-Object -First 10)
            byTotalFlights = ($squadronRankings | Sort-Object -Property totalFlights -Descending | Select-Object -First 10)
        }
        allSquadrons = $squadronRankings
        calculatedAt = (Get-Date -Format o)
    }

    $saved = Save-MetricsItem -Item $squadronRankingsMetric -ConnectionString $cosmosConnectionString -Database $cosmosDatabase
    Write-Log "$logPrefix   ✅ Squadron rankings saved"

    # METRIC 9: New Cadet Time-to-First-Flight
    Write-Log "$logPrefix   Calculating new cadet time-to-first-flight..."

    $timeToFirstFlightData = @()
    foreach ($user in $allUsers) {
        if ($user.employeeType -eq 'Cadet' -and $user.employeeId -and $user.createdDateTime) {
            $joinDate = [DateTime]::Parse($user.createdDateTime)
            $firstFlight = $allFlights | Where-Object { $_.CAPID -eq $user.employeeId } |
                Sort-Object { [DateTime]::Parse($_.FirstFlight) } |
                Select-Object -First 1

            if ($firstFlight) {
                $firstFlightDate = [DateTime]::Parse($firstFlight.FirstFlight)
                $daysToFirstFlight = ($firstFlightDate - $joinDate).Days

                if ($daysToFirstFlight -ge 0) {
                    $timeToFirstFlightData += [PSCustomObject]@{
                        capid = $user.employeeId
                        joinDate = $joinDate.ToString('yyyy-MM-dd')
                        firstFlightDate = $firstFlightDate.ToString('yyyy-MM-dd')
                        daysToFirstFlight = $daysToFirstFlight
                    }
                }
            }
        }
    }

    $avgDays = if ($timeToFirstFlightData.Count -gt 0) {
        [math]::Round(($timeToFirstFlightData | Measure-Object -Property daysToFirstFlight -Average).Average, 1)
    } else { 0 }

    $medianDays = if ($timeToFirstFlightData.Count -gt 0) {
        $sorted = $timeToFirstFlightData | Sort-Object -Property daysToFirstFlight
        $mid = [math]::Floor($sorted.Count / 2)
        $sorted[$mid].daysToFirstFlight
    } else { 0 }

    $timeToFirstFlightMetric = [PSCustomObject]@{
        id = "time-to-first-flight-$($now.ToString('yyyy-MM'))"
        metricType = "time-to-first-flight"
        calculatedDate = $now.ToString('yyyy-MM-dd')
        averageDays = $avgDays
        medianDays = $medianDays
        totalCadets = $timeToFirstFlightData.Count
        distribution = @{
            "0-30" = ($timeToFirstFlightData | Where-Object { $_.daysToFirstFlight -le 30 }).Count
            "31-60" = ($timeToFirstFlightData | Where-Object { $_.daysToFirstFlight -gt 30 -and $_.daysToFirstFlight -le 60 }).Count
            "61-90" = ($timeToFirstFlightData | Where-Object { $_.daysToFirstFlight -gt 60 -and $_.daysToFirstFlight -le 90 }).Count
            "91-180" = ($timeToFirstFlightData | Where-Object { $_.daysToFirstFlight -gt 90 -and $_.daysToFirstFlight -le 180 }).Count
            "181+" = ($timeToFirstFlightData | Where-Object { $_.daysToFirstFlight -gt 180 }).Count
        }
        calculatedAt = (Get-Date -Format o)
    }

    $saved = Save-MetricsItem -Item $timeToFirstFlightMetric -ConnectionString $cosmosConnectionString -Database $cosmosDatabase
    Write-Log "$logPrefix   ✅ Time-to-first-flight metrics saved (avg: $avgDays days)"

    # METRIC 11: Squadron Leaderboards
    Write-Log "$logPrefix   Calculating squadron leaderboards..."

    # Most flights this month
    $monthFlightsBySquadron = @{}
    foreach ($flight in $previousMonthFlights) {
        $squadron = $capidToSquadron[$flight.CAPID]
        if ($squadron) {
            if (-not $monthFlightsBySquadron.ContainsKey($squadron)) {
                $monthFlightsBySquadron[$squadron] = 0
            }
            $monthFlightsBySquadron[$squadron]++
        }
    }

    $topFlightsThisMonth = $monthFlightsBySquadron.GetEnumerator() |
        Sort-Object -Property Value -Descending |
        Select-Object -First 5 |
        ForEach-Object {
            [PSCustomObject]@{ squadron = $_.Key; flights = $_.Value }
        }

    # Highest completion rate (5-for-5)
    $topCompletion = @()
    foreach ($squadron in $activeCadetsBySquadron.Keys) {
        if ($syllabusCompletionMetric.squadrons.ContainsKey($squadron)) {
            $topCompletion += [PSCustomObject]@{
                squadron = $squadron
                completionRate = $syllabusCompletionMetric.squadrons[$squadron]["5for5"].percentage
                completers = $syllabusCompletionMetric.squadrons[$squadron]["5for5"].completed
            }
        }
    }
    $topCompletion = $topCompletion | Sort-Object -Property completionRate -Descending | Select-Object -First 5

    # Most new flyers (cadets with first flight in previous month)
    $newFlyersBySquadron = @{}
    foreach ($flight in $previousMonthFlights) {
        $squadron = $capidToSquadron[$flight.CAPID]
        if ($squadron -and $flight.Syllabus -eq "6") {
            if (-not $newFlyersBySquadron.ContainsKey($squadron)) {
                $newFlyersBySquadron[$squadron] = 0
            }
            $newFlyersBySquadron[$squadron]++
        }
    }

    $topNewFlyers = $newFlyersBySquadron.GetEnumerator() |
        Sort-Object -Property Value -Descending |
        Select-Object -First 5 |
        ForEach-Object {
            [PSCustomObject]@{ squadron = $_.Key; newFlyers = $_.Value }
        }

    $leaderboardMetric = [PSCustomObject]@{
        id = "squadron-leaderboards-$($previousMonthStart.ToString('yyyy-MM'))"
        metricType = "squadron-leaderboards"
        period = $previousMonthName
        calculatedDate = $now.ToString('yyyy-MM-dd')
        leaderboards = @{
            mostFlightsThisMonth = $topFlightsThisMonth
            highestCompletionRate = $topCompletion
            mostNewFlyers = $topNewFlyers
            topEngagement = ($squadronRankings | Sort-Object -Property engagementRate -Descending | Select-Object -First 5 |
                ForEach-Object { [PSCustomObject]@{ squadron = $_.squadron; engagementRate = $_.engagementRate } })
        }
        calculatedAt = (Get-Date -Format o)
    }

    $saved = Save-MetricsItem -Item $leaderboardMetric -ConnectionString $cosmosConnectionString -Database $cosmosDatabase
    Write-Log "$logPrefix   ✅ Squadron leaderboards saved"

    # SQUADRON PARTICIPATION DETAILS REPORT
    Write-Log "$logPrefix   Calculating squadron participation details report..."

    $squadronDetailsReport = @()

    # Calculate time to first flight by squadron
    $timeToFirstFlightBySquadron = @{}
    foreach ($user in $allUsers) {
        if ($user.employeeType -eq 'Cadet' -and $user.employeeId -and $user.createdDateTime -and $user.companyName) {
            if ($user.companyName -match 'CO-(.+)') {
                $squadron = $matches[1]
                $joinDate = [DateTime]::Parse($user.createdDateTime)
                $firstFlight = $allFlights | Where-Object { $_.CAPID -eq $user.employeeId } |
                    Sort-Object { [DateTime]::Parse($_.FirstFlight) } |
                    Select-Object -First 1

                if ($firstFlight) {
                    $firstFlightDate = [DateTime]::Parse($firstFlight.FirstFlight)
                    $daysToFirstFlight = ($firstFlightDate - $joinDate).Days

                    if ($daysToFirstFlight -ge 0) {
                        if (-not $timeToFirstFlightBySquadron.ContainsKey($squadron)) {
                            $timeToFirstFlightBySquadron[$squadron] = @()
                        }
                        $timeToFirstFlightBySquadron[$squadron] += $daysToFirstFlight
                    }
                }
            }
        }
    }

    # Get all cadets with at least one flight (calculated once for efficiency)
    $allCadetsWithFlights = $allFlights | Select-Object -ExpandProperty CAPID -Unique

    # Build the detailed report for each squadron
    foreach ($squadron in $activeCadetsBySquadron.Keys) {
        $sqData = $activeCadetsBySquadron[$squadron]
        $sqCadets = $sqData.CadetList

        # Total cadets in squadron (from Azure AD)
        $totalCadets = $sqData.TotalCadets

        # Cadets with at least one flight
        $sqCadetsWithFlights = ($sqCadets | Where-Object { $allCadetsWithFlights -contains $_ }).Count

        # Participation rate
        $participationRate = if ($totalCadets -gt 0) {
            [math]::Round(($sqCadetsWithFlights / $totalCadets) * 100, 1)
        } else { 0 }

        # Total flights for this squadron in FY
        $totalFlights = if ($fyBySquadron.ContainsKey($squadron)) { $fyBySquadron[$squadron].Total } else { 0 }

        # 5-for-5 completers
        $fiveForFiveCompleters = if ($syllabusCompletionMetric.squadrons.ContainsKey($squadron)) {
            $syllabusCompletionMetric.squadrons[$squadron]["5for5"].completed
        } else { 0 }

        # 5-for-5 completion rate
        $completionRate = if ($totalCadets -gt 0) {
            [math]::Round(($fiveForFiveCompleters / $totalCadets) * 100, 1)
        } else { 0 }

        # Average time to first flight for this squadron
        $avgTimeToFirstFlight = if ($timeToFirstFlightBySquadron.ContainsKey($squadron) -and $timeToFirstFlightBySquadron[$squadron].Count -gt 0) {
            [math]::Round(($timeToFirstFlightBySquadron[$squadron] | Measure-Object -Average).Average, 1)
        } else { 0 }

        $squadronDetailsReport += [PSCustomObject]@{
            squadron = $squadron
            totalCadets = $totalCadets
            cadetsWithFlights = $sqCadetsWithFlights
            participationRate = $participationRate
            totalFlights = $totalFlights
            fiveForFiveCompleters = $fiveForFiveCompleters
            completionRate = $completionRate
            avgTimeToFirstFlight = $avgTimeToFirstFlight
        }
    }

    # Sort by squadron number
    $squadronDetailsReport = $squadronDetailsReport | Sort-Object -Property squadron

    # Save the squadron participation details report
    $squadronDetailsMetric = [PSCustomObject]@{
        id = "squadron-participation-details-$($now.ToString('yyyy-MM'))"
        metricType = "squadron-participation-details"
        fiscalYear = $fyStart.Year + 1
        calculatedDate = $now.ToString('yyyy-MM-dd')
        squadrons = $squadronDetailsReport
        wingTotals = @{
            totalCadets = ($squadronDetailsReport | Measure-Object -Property totalCadets -Sum).Sum
            cadetsWithFlights = ($squadronDetailsReport | Measure-Object -Property cadetsWithFlights -Sum).Sum
            participationRate = if (($squadronDetailsReport | Measure-Object -Property totalCadets -Sum).Sum -gt 0) {
                [math]::Round((($squadronDetailsReport | Measure-Object -Property cadetsWithFlights -Sum).Sum / ($squadronDetailsReport | Measure-Object -Property totalCadets -Sum).Sum) * 100, 1)
            } else { 0 }
            totalFlights = ($squadronDetailsReport | Measure-Object -Property totalFlights -Sum).Sum
            fiveForFiveCompleters = ($squadronDetailsReport | Measure-Object -Property fiveForFiveCompleters -Sum).Sum
            completionRate = if (($squadronDetailsReport | Measure-Object -Property totalCadets -Sum).Sum -gt 0) {
                [math]::Round((($squadronDetailsReport | Measure-Object -Property fiveForFiveCompleters -Sum).Sum / ($squadronDetailsReport | Measure-Object -Property totalCadets -Sum).Sum) * 100, 1)
            } else { 0 }
            avgTimeToFirstFlight = if ($timeToFirstFlightData.Count -gt 0) {
                [math]::Round(($timeToFirstFlightData | Measure-Object -Property daysToFirstFlight -Average).Average, 1)
            } else { 0 }
        }
        calculatedAt = (Get-Date -Format o)
    }

    $saved = Save-MetricsItem -Item $squadronDetailsMetric -ConnectionString $cosmosConnectionString -Database $cosmosDatabase
    Write-Log "$logPrefix   ✅ Squadron participation details report saved"

    Write-Log "$logPrefix ✅ All additional metrics calculated and saved"
    Write-Log "$logPrefix"

    # ============================================================================
    # OFLIGHT PRIORITY CALCULATION
    # ============================================================================
    Write-Log "$logPrefix Calculating OFlight Priority scores..."

    # Helper functions for priority calculation
    function Get-MonthsUntil18 {
        param([AllowNull()][Nullable[datetime]]$DOB, [datetime]$AsOf)
        if ($null -eq $DOB -or -not $DOB.HasValue) { return 999 }
        $eighteenth = $DOB.Value.AddYears(18)
        $days = ($eighteenth - $AsOf).TotalDays
        if ($days -le 0) { return 0 }
        return [int][math]::Floor($days / 30.44)
    }

    function Get-AgeUrgencyPoints {
        param([int]$MonthsUntil18)
        switch ($MonthsUntil18) {
            { $_ -le 3 }  { return 300 }
            { $_ -le 6 }  { return 200 }
            { $_ -le 12 } { return 100 }
            { $_ -le 18 } { return 50 }
            default       { return 0  }
        }
    }

    function Get-ProgressionPoints {
        param([int]$FlightsCompleted)
        $points = (5 - $FlightsCompleted) * 6
        if ($points -lt 0) { $points = 0 }
        if ($points -gt 30) { $points = 30 }
        return $points
    }

    function Get-FirstFlightUrgency {
        param([int]$FlightsCompleted, [int]$DaysSinceJoin, [int]$FirstFlightDaysThreshold = 60)
        if ($FlightsCompleted -ne 0) { return 0 }
        # 1 point per day since joining (no cap)
        return $DaysSinceJoin
    }

    function Get-SinceLastFlightPoints {
        param([int]$FlightsCompleted, [nullable[datetime]]$LastFlightDate, [datetime]$AsOf)
        if ($FlightsCompleted -le 0 -or -not $LastFlightDate) { return 0 }
        $days = ($AsOf - $LastFlightDate).TotalDays
        # 1 point per day since last flight (no cap)
        return [int][math]::Floor($days)
    }

    function Get-Tier {
        param([int]$FlightsCompleted,[int]$DaysSinceJoin,[int]$DaysSinceLast,[int]$MonthsUntil18,[int]$FirstFlightDaysThreshold = 60)
        <#
        Tiering rules (evaluated in this order):

        1) COMPLETED: `FlightsCompleted >= 5` OR `MonthsUntil18 == 0` (18 years old or older)

        2) Critical: any of
           - `FlightsCompleted == 0` AND `DaysSinceJoin >= 180` (cadets rarely fly within first 60 days due to uniform requirement)
           - `FlightsCompleted > 0` AND `DaysSinceLast >= 240`
           - `MonthsUntil18 <= 3` AND `FlightsCompleted < 5`

        3) High: any of
           - `FlightsCompleted == 0` AND `DaysSinceJoin >= 120` AND `DaysSinceJoin < 180`
           - `FlightsCompleted >= 1` AND `DaysSinceLast >= 90` AND `DaysSinceLast < 240`
           - `MonthsUntil18 <= 12` AND `MonthsUntil18 > 3` AND `FlightsCompleted < 5`

        4) Medium: any of
           - `FlightsCompleted == 0` AND `DaysSinceJoin >= 90` AND `DaysSinceJoin < 120`
           - `FlightsCompleted >= 1` AND `DaysSinceLast >= 30` AND `DaysSinceLast < 90`
           - `MonthsUntil18 > 12` AND `MonthsUntil18 <= 18` AND `FlightsCompleted < 5`

        5) Low: default catch-all for remaining cadets (e.g., recent joiners or recent flights)

        Target distribution guidance (informational):
        - Critical: ~5-10%
        - High: ~15-25%
        - Medium: ~35-45%
        - Low: ~10-20%
        - COMPLETED: varies
        #>

        if ($FlightsCompleted -ge 5 -or $MonthsUntil18 -eq 0) { return 'COMPLETED' }

        if ( ($FlightsCompleted -eq 0 -and $DaysSinceJoin -ge 180) -or
             ($FlightsCompleted -gt 0 -and $DaysSinceLast -ge 240) -or
             ($MonthsUntil18 -le 3 -and $FlightsCompleted -lt 5) ) {
            return 'Critical'
        }

        if ( ($FlightsCompleted -eq 0 -and $DaysSinceJoin -ge 120 -and $DaysSinceJoin -lt 180) -or
             ($FlightsCompleted -ge 1 -and $DaysSinceLast -ge 90 -and $DaysSinceLast -lt 240) -or
             ($MonthsUntil18 -le 12 -and $MonthsUntil18 -gt 3 -and $FlightsCompleted -lt 5) ) {
            return 'High'
        }

        if ( ($FlightsCompleted -eq 0 -and $DaysSinceJoin -ge 90 -and $DaysSinceJoin -lt 120) -or
             ($FlightsCompleted -ge 1 -and $DaysSinceLast -ge 30 -and $DaysSinceLast -lt 90) -or
             ($MonthsUntil18 -gt 12 -and $MonthsUntil18 -le 18 -and $FlightsCompleted -lt 5) ) {
            return 'Medium'
        }

        return 'Low'
    }

    # Get all cadets with their flight data
    $priorityCadets = @()
    $AsOfDate = Get-Date

    # Build a lookup of flights by CAPID
    $flightsByCapid = @{}
    foreach ($flight in $allFlights) {
        $capid = $flight.CAPID
        if (-not $flightsByCapid.ContainsKey($capid)) {
            $flightsByCapid[$capid] = @()
        }
        $flightsByCapid[$capid] += $flight
    }

    foreach ($user in $allUsers) {
        if ($user.employeeType -eq 'Cadet' -and $user.employeeId) {
            $capid = $user.employeeId
            $squadron = if ($user.companyName -match 'CO-(.+)') { $matches[1] } else { $null }

            # Get flight data for this cadet
            $cadetFlights = if ($flightsByCapid.ContainsKey($capid)) { $flightsByCapid[$capid] } else { @() }
            $flightsCompleted = [math]::Min($cadetFlights.Count, 5)

            # Find last flight date
            $lastFlightDate = $null
            if ($cadetFlights.Count -gt 0) {
                $sortedFlights = $cadetFlights | Sort-Object {
                    try { [DateTime]::Parse($_.FirstFlight) } catch { [DateTime]::MinValue }
                } -Descending
                try {
                    $lastFlightDate = [DateTime]::Parse($sortedFlights[0].FirstFlight)
                } catch { }
            }

            # Parse dates
            $dob = $null
            $joinedDate = $null
            try {
                if ($user.createdDateTime) { $joinedDate = [DateTime]::Parse($user.createdDateTime) }
            } catch { }

            # DOB is not available in user object, but we can check onPremisesExtensionAttributes or other fields
            # For now, we'll skip DOB-based calculations unless it's available elsewhere
            # You may need to load from Member.txt if DOB is required

            $daysSinceJoin = if ($joinedDate) { [int][math]::Floor(($AsOfDate - $joinedDate).TotalDays) } else { $null }
            $daysSinceLast = if ($lastFlightDate) { [int][math]::Floor(($AsOfDate - $lastFlightDate).TotalDays) } else { $null }

            $monthsUntil18 = Get-MonthsUntil18 -DOB $dob -AsOf $AsOfDate
            $ageYears = if ($dob) { [int][math]::Floor((($AsOfDate - $dob).TotalDays) / 365.25) } else { $null }

            # Calculate priority components
            $A = if ($daysSinceJoin -ne $null) { Get-FirstFlightUrgency -FlightsCompleted $flightsCompleted -DaysSinceJoin $daysSinceJoin } else { 0 }
            $B = Get-SinceLastFlightPoints -FlightsCompleted $flightsCompleted -LastFlightDate $lastFlightDate -AsOf $AsOfDate
            $C = Get-ProgressionPoints -FlightsCompleted $flightsCompleted
            $D = Get-AgeUrgencyPoints -MonthsUntil18 $monthsUntil18

            # Cadets who completed all 5 flights OR are 18+ years old are marked COMPLETED with priority 0
            if ($flightsCompleted -ge 5 -or $monthsUntil18 -eq 0) {
                $priority = 0
                $tier = 'COMPLETED'
                $A = 0
                $B = 0
                $C = 0
                $D = 0
            } else {
                $priority = [math]::Round(($A + $B + $C + $D), 2)
                $tier = Get-Tier -FlightsCompleted $flightsCompleted -DaysSinceJoin ($daysSinceJoin ?? 0) -DaysSinceLast ($daysSinceLast ?? 0) -MonthsUntil18 $monthsUntil18
            }

            $nextFlight = if ($flightsCompleted -ge 5) { 5 } else { $flightsCompleted + 1 }

            # Extract name from displayName
            $firstName = ""
            $lastName = ""
            if ($user.displayName) {
                $nameParts = $user.displayName -split ','
                if ($nameParts.Count -ge 2) {
                    $lastName = $nameParts[0].Trim()
                    $firstName = ($nameParts[1] -replace '\s+\w+$', '').Trim()  # Remove rank at end
                } else {
                    $fullName = $user.displayName
                    $nameComponents = $fullName -split ' '
                    if ($nameComponents.Count -ge 2) {
                        $firstName = $nameComponents[0]
                        $lastName = $nameComponents[-2]  # Second to last (before rank)
                    }
                }
            }

            $priorityCadets += [pscustomobject]@{
                CAPID                  = $capid
                LastName               = $lastName
                FirstName              = $firstName
                Email                  = $user.mail
                Squadron               = $squadron
                DOB                    = $dob
                AgeYears               = $ageYears
                MonthsUntil18          = $monthsUntil18
                JoinedDate             = $joinedDate
                DaysSinceJoin          = $daysSinceJoin
                LastFlightDate         = $lastFlightDate
                DaysSinceLastFlight    = $daysSinceLast
                FlightsCompleted       = $flightsCompleted
                NextFlightNumber       = $nextFlight
                A_FirstFlightUrgency   = $A
                B_SinceLastFlight      = $B
                C_ProgressionEquity    = $C
                D_AgeUrgency           = $D
                PriorityScore          = $priority
                Tier                   = $tier
            }
        }
    }

    # Sort by priority score
    $prioritized = $priorityCadets | Sort-Object `
        @{Expression='PriorityScore';Descending=$true},
        @{Expression='FlightsCompleted';Ascending=$true},
        @{Expression='DaysSinceLastFlight';Descending=$true}

    Write-Log "$logPrefix   Calculated priority for $($prioritized.Count) cadets"

    # Group by squadron for squadron-level metrics
    $squadronPriorityMetrics = @{}
    foreach ($cadet in $prioritized) {
        $sq = $cadet.Squadron
        if ($sq) {
            if (-not $squadronPriorityMetrics.ContainsKey($sq)) {
                $squadronPriorityMetrics[$sq] = @{
                    totalCadets = 0
                    byTier = @{
                        Critical = 0
                        High = 0
                        Medium = 0
                        Low = 0
                        COMPLETED = 0
                    }
                    avgPriorityScore = 0
                    cadets = @()
                }
            }

            $squadronPriorityMetrics[$sq].totalCadets++
            $squadronPriorityMetrics[$sq].byTier[$cadet.Tier]++
            $squadronPriorityMetrics[$sq].cadets += [PSCustomObject]@{
                capid = $cadet.CAPID
                lastName = $cadet.LastName
                firstName = $cadet.FirstName
                email = $cadet.Email
                flightsCompleted = $cadet.FlightsCompleted
                nextFlightNumber = $cadet.NextFlightNumber
                priorityScore = $cadet.PriorityScore
                tier = $cadet.Tier
                monthsUntil18 = $cadet.MonthsUntil18
                daysSinceJoin = $cadet.DaysSinceJoin
                daysSinceLastFlight = $cadet.DaysSinceLastFlight
            }
        }
    }

    # Calculate average priority scores
    foreach ($sq in $squadronPriorityMetrics.Keys) {
        $scores = $squadronPriorityMetrics[$sq].cadets | ForEach-Object { $_.priorityScore }
        $squadronPriorityMetrics[$sq].avgPriorityScore = if ($scores.Count -gt 0) {
            [math]::Round(($scores | Measure-Object -Average).Average, 2)
        } else { 0 }
    }

    # Create the priority metrics document
    $priorityMetric = [PSCustomObject]@{
        id = "oflight-priority-$($now.ToString('yyyy-MM-dd'))"
        metricType = "oflight-priority"
        calculatedDate = $now.ToString('yyyy-MM-dd')
        totalCadets = $prioritized.Count
        byTier = @{ 
            Critical = ($prioritized | Where-Object { $_.Tier -eq 'Critical' }).Count
            High = ($prioritized | Where-Object { $_.Tier -eq 'High' }).Count
            Medium = ($prioritized | Where-Object { $_.Tier -eq 'Medium' }).Count
            Low = ($prioritized | Where-Object { $_.Tier -eq 'Low' }).Count
            COMPLETED = ($prioritized | Where-Object { $_.Tier -eq 'COMPLETED' }).Count
        }
        avgPriorityScore = if ($prioritized.Count -gt 0) {
            [math]::Round(($prioritized | Measure-Object -Property PriorityScore -Average).Average, 2)
        } else { 0 }
        squadrons = $squadronPriorityMetrics
        topPriority = ($prioritized | Select-Object -First 20 | ForEach-Object {
            [PSCustomObject]@{
                capid = $_.CAPID
                squadron = $_.Squadron
                lastName = $_.LastName
                firstName = $_.FirstName
                email = $_.Email
                priorityScore = $_.PriorityScore
                tier = $_.Tier
                flightsCompleted = $_.FlightsCompleted
                daysSinceJoin = $_.DaysSinceJoin
            }
        })
        calculatedAt = (Get-Date -Format o)
    }

    # Save priority metrics
    $saved = Save-MetricsItem -Item $priorityMetric -ConnectionString $cosmosConnectionString -Database $cosmosDatabase
    if ($saved) {
        Write-Log "$logPrefix   ✅ OFlight Priority metrics saved"
        Write-Log "$logPrefix      Total cadets: $($prioritized.Count)"

        # Tier distribution counts and percentages
        $total = $priorityMetric.totalCadets
        if ($total -gt 0) {
            $critPct = [math]::Round(($priorityMetric.byTier.Critical / $total) * 100, 1)
            $highPct = [math]::Round(($priorityMetric.byTier.High / $total) * 100, 1)
            $medPct = [math]::Round(($priorityMetric.byTier.Medium / $total) * 100, 1)
            $lowPct = [math]::Round(($priorityMetric.byTier.Low / $total) * 100, 1)
            $compPct = [math]::Round(($priorityMetric.byTier.COMPLETED / $total) * 100, 1)
        } else {
            $critPct = $highPct = $medPct = $lowPct = $compPct = 0
        }

        Write-Log "$logPrefix      Critical: $($priorityMetric.byTier.Critical) ($critPct%) | High: $($priorityMetric.byTier.High) ($highPct%) | Medium: $($priorityMetric.byTier.Medium) ($medPct%) | Low: $($priorityMetric.byTier.Low) ($lowPct%) | COMPLETED: $($priorityMetric.byTier.COMPLETED) ($compPct%)"
        Write-Log "$logPrefix      Avg Priority Score: $($priorityMetric.avgPriorityScore)"

        # Zero-flight buckets for operational visibility
        $zero0_29 = ($prioritized | Where-Object { $_.FlightsCompleted -eq 0 -and $_.DaysSinceJoin -ne $null -and $_.DaysSinceJoin -lt 30 }).Count
        $zero30_59 = ($prioritized | Where-Object { $_.FlightsCompleted -eq 0 -and $_.DaysSinceJoin -ne $null -and $_.DaysSinceJoin -ge 30 -and $_.DaysSinceJoin -lt 60 }).Count
        $zero60_120 = ($prioritized | Where-Object { $_.FlightsCompleted -eq 0 -and $_.DaysSinceJoin -ne $null -and $_.DaysSinceJoin -ge 60 -and $_.DaysSinceJoin -le 120 }).Count
        $zero121plus = ($prioritized | Where-Object { $_.FlightsCompleted -eq 0 -and $_.DaysSinceJoin -ne $null -and $_.DaysSinceJoin -gt 120 }).Count

        Write-Log "$logPrefix      Zero-flight by days-since-join: 0-29: $zero0_29 | 30-59: $zero30_59 | 60-120: $zero60_120 | 121+: $zero121plus"
    } else {
        Write-Log "$logPrefix   ❌ Failed to save OFlight Priority metrics"
    }

    Write-Log "$logPrefix"

    # Log summary
    Write-Log "$logPrefix"
    Write-Log "$logPrefix ==================== METRICS SUMMARY ===================="
    Write-Log "$logPrefix"
    Write-Log "$logPrefix Previous Month ($previousMonthName):"
    Write-Log "$logPrefix   Total Flights: $($previousMonthFlights.Count)"
    Write-Log "$logPrefix   Unique Cadets: $(($previousMonthFlights | Select-Object -ExpandProperty CAPID -Unique).Count)"
    Write-Log "$logPrefix   Squadrons: $($previousMonthBySquadron.Keys.Count)"
    foreach ($squadron in ($previousMonthBySquadron.Keys | Sort-Object)) {
        Write-Log "$logPrefix     $squadron : $($previousMonthBySquadron[$squadron].Total) flights, $($previousMonthBySquadron[$squadron].UniqueCadets.Count) cadets"
    }
    Write-Log "$logPrefix"
    Write-Log "$logPrefix Fiscal Year ($fyName):"
    Write-Log "$logPrefix   Total Flights: $($fyFlights.Count)"
    Write-Log "$logPrefix   Unique Cadets: $(($fyFlights | Select-Object -ExpandProperty CAPID -Unique).Count)"
    Write-Log "$logPrefix   Squadrons: $($fyBySquadron.Keys.Count)"
    foreach ($squadron in ($fyBySquadron.Keys | Sort-Object)) {
        Write-Log "$logPrefix     $squadron : $($fyBySquadron[$squadron].Total) flights, $($fyBySquadron[$squadron].UniqueCadets.Count) cadets"
    }
    Write-Log "$logPrefix"
    Write-Log "$logPrefix ========================================================="
    Write-Log "$logPrefix"

    Write-Log "$logPrefix ✅ OFlightMetrics function completed successfully"

} catch {
    Write-Log "$logPrefix ❌ Error in OFlightMetrics function: $($_.Exception.Message)"
    Write-Log "$logPrefix Stack trace: $($_.ScriptStackTrace)"
    throw
}
