<#
.SYNOPSIS
  CAP Cadet Orientation Flight Prioritizer (Members + OFlights)

.DESCRIPTION
  Reads Member.txt and OFlight.txt, merges by CAPID, computes a priority score for O-Flights,
  and optionally builds a flight-day schedule with allocation and per-squadron caps.
  Can save results to Cosmos DB Metrics container.

.PARAMETER MemberPath
  Path to Member.txt (default: $env:HOME\data\CAPWatch\Member.txt).

.PARAMETER OFlightsPath
  Path to OFlight.txt (default: $env:HOME\data\CAPWatch\OFlight.txt).

.PARAMETER OutputCsv
  Path for the full prioritized list (CSV).

.PARAMETER OutputScheduleCsv
  Optional path for a flight-day schedule CSV.

.PARAMETER AsOf
  Reference date (default: today).

.PARAMETER TotalSlots
  If set with OutputScheduleCsv, builds a schedule with N slots.

.PARAMETER AllocationFirst
  Fraction (0–1) for first flights (default: 0.40).

.PARAMETER AllocationProgress
  Fraction (0–1) for progressing flights #2–#4 (default: 0.40).

.PARAMETER AllocationAgeCritical
  Fraction (0–1) for age-critical cadets (default: 0.20).

.PARAMETER MaxPerSquadron
  Optional cap per squadron in the schedule.

.PARAMETER FirstFlightDaysThreshold
  Days target for first flight after join (default: 60).

.PARAMETER SaveToCosmosDb
  If set, saves priority results to Cosmos DB Metrics container.

.PARAMETER ConnectionString
  Cosmos DB connection string (default: $env:CosmosDbConnectionString).

.PARAMETER Database
  Cosmos DB database name (default: $env:CosmosDbDatabase).

.NOTES
  Author: Michael Schulte, Capt
  Based on algorithm provided for CAP Orientation Flight Prioritization
#>

[CmdletBinding()]
param(
    [string]$MemberPath = "$($env:HOME)\data\CAPWatch\Member.txt",
    [string]$OFlightsPath = "$($env:HOME)\data\CAPWatch\OFlight.txt",
    [string]$OutputCsv = ".\OFlightPriority.csv",

    [string]$OutputScheduleCsv,
    [datetime]$AsOf = (Get-Date),
    [int]$TotalSlots,

    [double]$AllocationFirst = 0.40,
    [double]$AllocationProgress = 0.40,
    [double]$AllocationAgeCritical = 0.20,

    [int]$MaxPerSquadron,
    [int]$FirstFlightDaysThreshold = 60,

    [switch]$SaveToCosmosDb,
    [string]$ConnectionString = $env:CosmosDbConnectionString,
    [string]$Database = $env:CosmosDbDatabase
)

# ---------- Utilities ----------

function Get-Delimiter {
    param([Parameter(Mandatory=$true)][string]$Path)
    $sample = Get-Content -Path $Path -TotalCount 1
    $candidates = @(",","`t","|",";")
    $best = ","; $bestCount = -1
    foreach ($d in $candidates) {
        $count = ($sample -split [regex]::Escape($d)).Count - 1
        if ($count -gt $bestCount) { $best = $d; $bestCount = $count }
    }
    return $best
}

function Import-Table {
    param([Parameter(Mandatory=$true)][string]$Path)
    if (-not (Test-Path $Path)) { throw "File not found: $Path" }
    $delim = Get-Delimiter -Path $Path
    return Import-Csv -Path $Path -Delimiter $delim
}

function Get-MonthsUntil18 {
    param([AllowNull()][Nullable[datetime]]$DOB, [datetime]$AsOf)
    if ($null -eq $DOB -or -not $DOB.HasValue) { return 999 }
    $eighteenth = $DOB.Value.AddYears(18)
    $days = ($eighteenth - $AsOf).TotalDays
    if ($days -le 0) { return 0 }
    return [int][math]::Floor($days / 30.44) # average month length
}

function Get-AgeUrgencyPoints {
    param([int]$MonthsUntil18)
    switch ($MonthsUntil18) {
        { $_ -le 3 }  { return 40 }
        { $_ -le 6 }  { return 30 }
        { $_ -le 12 } { return 20 }
        { $_ -le 18 } { return 10 }
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
    param([int]$FlightsCompleted, [int]$DaysSinceJoin, [int]$FirstFlightDaysThreshold)
    if ($FlightsCompleted -ne 0) { return 0 }
    $ratio = [double]$DaysSinceJoin / [double]$FirstFlightDaysThreshold
    $score = [math]::Min(1.0, [math]::Max(0.0, $ratio)) * 100
    return [math]::Round($score, 2)
}

function Get-SinceLastFlightPoints {
    param([int]$FlightsCompleted, [nullable[datetime]]$LastFlightDate, [datetime]$AsOf)
    if ($FlightsCompleted -le 0 -or -not $LastFlightDate) { return 0 }
    $days = ($AsOf - $LastFlightDate).TotalDays
    $score = [math]::Min(40, ($days / 30.0) * 10)   # ~10 pts / month
    return [math]::Round($score, 2)
}

function Get-Tier {
    param([int]$FlightsCompleted,[int]$DaysSinceJoin,[int]$DaysSinceLast,[int]$MonthsUntil18,[int]$FirstFlightDaysThreshold)
    if ( ($FlightsCompleted -eq 0 -and $DaysSinceJoin -gt $FirstFlightDaysThreshold) -or
         ($FlightsCompleted -gt 0 -and $DaysSinceLast -ge 180) -or
         ($MonthsUntil18 -le 6 -and $FlightsCompleted -lt 5) ) {
        return 'Critical'
    }
    elseif ( ($FlightsCompleted -eq 0 -and $DaysSinceJoin -ge [math]::Floor($FirstFlightDaysThreshold * 0.5)) -or
             ($FlightsCompleted -ge 1 -and $DaysSinceLast -ge 90) -or
             ($MonthsUntil18 -le 12 -and $FlightsCompleted -lt 5) ) {
        return 'High'
    }
    elseif ($FlightsCompleted -lt 5) { return 'Medium' }
    else { return 'Low' }
}

function Add-Jitter {
    param([string]$Key)
    $sha1 = [System.Security.Cryptography.SHA1]::Create()
    $hash = $sha1.ComputeHash([Text.Encoding]::UTF8.GetBytes($Key))
    $u32  = [BitConverter]::ToUInt32($hash,0)
    return ($u32 % 1000) / 100000.0 # 0 to 0.00999
}

function Save-MetricsItem {
    param (
        [Parameter(Mandatory=$true)]
        [object]$Item,
        [string]$ConnectionString,
        [string]$Database
    )

    try {
        # Ensure item has required id field
        if (-not $Item.id) {
            Write-Host "❌ Error: Metrics item must have an 'id' field" -ForegroundColor Red
            return $false
        }

        # Ensure item has metricType for partition key
        if (-not $Item.metricType) {
            Write-Host "❌ Error: Metrics item must have a 'metricType' field for partition key" -ForegroundColor Red
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
            Write-Host "❌ Failed to parse Cosmos DB connection string" -ForegroundColor Red
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

        Write-Host "✅ Successfully saved metrics to Cosmos DB: $($Item.id)" -ForegroundColor Green
        return $true

    } catch {
        $errorMessage = $_.Exception.Message
        Write-Host "❌ Failed to save metrics to Cosmos DB: $($Item.id). Error: $errorMessage" -ForegroundColor Red
        return $false
    }
}

# ---------- Load data ----------

Write-Host "Loading member data from $MemberPath..." -ForegroundColor Cyan
$memRaw = Import-Table -Path $MemberPath

Write-Host "Loading O-Flight data from $OFlightsPath..." -ForegroundColor Cyan
$oflRaw = Import-Table -Path $OFlightsPath

# Filter members: Only active cadets
$mem = $memRaw | Where-Object { $_.MbrStatus -eq "ACTIVE" -and $_.Type -eq "CADET" } | ForEach-Object {
    # Parse dates
    $dob = $null
    try {
        if ($_.DOB) { $dob = [datetime]::Parse($_.DOB) }
    } catch { }

    $joinedDt = $null
    try {
        if ($_.Joined) { $joinedDt = [datetime]::Parse($_.Joined) }
    } catch { }

    [pscustomobject]@{
        CAPID       = "$($_.CAPID)".Trim()
        LastName    = $_.NameLast
        FirstName   = $_.NameFirst
        DOB         = $dob
        JoinedDate  = $joinedDt
        Squadron    = $_.Unit
    }
}

Write-Host "Found $($mem.Count) active cadets in Member.txt" -ForegroundColor Green

# Process OFlights: Filter for syllabus 6-10, group by CAPID
$oflFiltered = $oflRaw | Where-Object { $_.Syllabus -in @("6", "7", "8", "9", "10") }

$ofl = $oflFiltered | Group-Object -Property CAPID | ForEach-Object {
    $flights = $_.Group
    $flightCount = $flights.Count

    # Find most recent flight date
    $lastFlight = $null
    foreach ($flight in $flights) {
        try {
            $fltDate = [datetime]::Parse($flight.FltDate)
            if (-not $lastFlight -or $fltDate -gt $lastFlight) {
                $lastFlight = $fltDate
            }
        } catch {
            # Skip invalid dates
            continue
        }
    }

    [pscustomobject]@{
        CAPID           = "$($_.Name)".Trim()
        FlightsCompleted= [math]::Min([math]::Max($flightCount,0),5)
        LastFlightDate  = $lastFlight
    }
}

Write-Host "Processed $($ofl.Count) cadets with O-Flight records (syllabus 6-10)" -ForegroundColor Green

# Build lookup for merges
$oflById = @{}
foreach ($r in $ofl) { $oflById[$r.CAPID] = $r }

# ---------- Merge + compute ----------
Write-Host "Computing priority scores..." -ForegroundColor Cyan

$enriched = foreach ($m in $mem) {
    $o = $oflById[$m.CAPID]
    $fc = if ($o) { $o.FlightsCompleted } else { 0 }
    $lfd = if ($o) { $o.LastFlightDate } else { $null }

    $daysSinceJoin = if ($m.JoinedDate) { [int][math]::Floor( ($AsOf - $m.JoinedDate).TotalDays ) } else { $null }
    $daysSinceLast = if ($lfd) { [int][math]::Floor( ($AsOf - $lfd).TotalDays ) } else { $null }

    $monthsUntil18 = Get-MonthsUntil18 -DOB $m.DOB -AsOf $AsOf
    $ageYears = if ($m.DOB) { [int][math]::Floor( (($AsOf - $m.DOB).TotalDays) / 365.25 ) } else { $null }

    # A/B/C/D components
    $A = if ($daysSinceJoin -ne $null) { Get-FirstFlightUrgency -FlightsCompleted $fc -DaysSinceJoin $daysSinceJoin -FirstFlightDaysThreshold $FirstFlightDaysThreshold } else { 0 }
    $B = Get-SinceLastFlightPoints -FlightsCompleted $fc -LastFlightDate $lfd -AsOf $AsOf
    $C = Get-ProgressionPoints -FlightsCompleted $fc
    $D = Get-AgeUrgencyPoints -MonthsUntil18 $monthsUntil18

    # Cadets who completed all 5 flights are marked COMPLETED with priority 0
    if ($fc -ge 5) {
        $priority = 0
        $tier = 'COMPLETED'
        $A = 0
        $B = 0
        $C = 0
        $D = 0
    } else {
        $priority = [math]::Round(($A + $B + $C + $D),2)
        $tier = Get-Tier -FlightsCompleted $fc -DaysSinceJoin ($daysSinceJoin ?? 0) -DaysSinceLast ($daysSinceLast ?? 0) -MonthsUntil18 $monthsUntil18 -FirstFlightDaysThreshold $FirstFlightDaysThreshold
    }

    $nextFlight = if ($fc -ge 5) { 5 } else { $fc + 1 }

    [pscustomobject]@{
        CAPID                  = $m.CAPID
        LastName               = $m.LastName
        FirstName              = $m.FirstName
        Squadron               = $m.Squadron
        DOB                    = $m.DOB
        AgeYears               = $ageYears
        MonthsUntil18          = $monthsUntil18
        JoinedDate             = $m.JoinedDate
        DaysSinceJoin          = $daysSinceJoin
        LastFlightDate         = $lfd
        DaysSinceLastFlight    = $daysSinceLast
        FlightsCompleted       = $fc
        NextFlightNumber       = $nextFlight
        A_FirstFlightUrgency   = $A
        B_SinceLastFlight      = $B
        C_ProgressionEquity    = $C
        D_AgeUrgency           = $D
        PriorityScore          = $priority
        Tier                   = $tier
        _TieFewestFlights      = (5 - $fc)       # higher => fewer flights
        _TieDaysSinceLast      = ($daysSinceLast ?? -1)
        _TieAgeYears           = ($ageYears ?? -1)
        _Jitter                = Add-Jitter -Key ("$($m.CAPID)$($m.LastName)$($m.FirstName)")
    }
}

# Sort by score + tie breakers
$prioritized = $enriched | Sort-Object `
    @{Expression='PriorityScore';Descending=$true},
    @{Expression='_TieFewestFlights';Descending=$true},
    @{Expression='_TieDaysSinceLast';Descending=$true},
    @{Expression='_TieAgeYears';Descending=$true},
    @{Expression='_Jitter';Descending=$true}

# Sort by Squadron for final output
$prioritizedBySquadron = $prioritized | Sort-Object Squadron, @{Expression='PriorityScore';Descending=$true}

# Export prioritized list
$prioritizedBySquadron | Select-Object CAPID,Squadron,LastName,FirstName,FlightsCompleted,NextFlightNumber,JoinedDate,DaysSinceJoin,LastFlightDate,DaysSinceLastFlight,DOB,AgeYears,MonthsUntil18,A_FirstFlightUrgency,B_SinceLastFlight,C_ProgressionEquity,D_AgeUrgency,PriorityScore,Tier |
    Export-Csv -NoTypeInformation -Path $OutputCsv -Encoding UTF8

Write-Host "✅ Wrote prioritized list (sorted by squadron) to $OutputCsv" -ForegroundColor Green
Write-Host "   Total cadets: $($prioritized.Count)" -ForegroundColor Gray

# Display summary by tier
$tierSummary = $prioritized | Group-Object Tier | Sort-Object Name
Write-Host "`nPriority Tier Summary:" -ForegroundColor Cyan
foreach ($tier in $tierSummary) {
    Write-Host "  $($tier.Name): $($tier.Count) cadets" -ForegroundColor White
}

# Display summary by squadron
$squadronSummary = $prioritized | Group-Object Squadron | Sort-Object Name
Write-Host "`nSquadron Summary:" -ForegroundColor Cyan
foreach ($sq in $squadronSummary) {
    $avgScore = [math]::Round(($sq.Group | Measure-Object -Property PriorityScore -Average).Average, 1)
    Write-Host "  Squadron $($sq.Name): $($sq.Count) cadets (avg priority: $avgScore)" -ForegroundColor White
}

# ---------- Optional: build schedule ----------
if ($TotalSlots -and $OutputScheduleCsv) {
    Write-Host "`nBuilding flight schedule with $TotalSlots slots..." -ForegroundColor Cyan

    # Compute category counts & fix rounding remainder
    $nFirst    = [int][math]::Round($TotalSlots * $AllocationFirst)
    $nProg     = [int][math]::Round($TotalSlots * $AllocationProgress)
    $nAgeCrit  = [int][math]::Round($TotalSlots * $AllocationAgeCritical)

    $sum = $nFirst + $nProg + $nAgeCrit
    if ($sum -ne $TotalSlots) {
        $diff = $TotalSlots - $sum
        while ($diff -ne 0) {
            if ($diff -gt 0) {
                if ($nAgeCrit -lt $TotalSlots) { $nAgeCrit++ }
                elseif ($nFirst -lt $TotalSlots) { $nFirst++ }
                else { $nProg++ }
                $diff--
            } else {
                if ($nProg -gt 0) { $nProg-- }
                elseif ($nFirst -gt 0) { $nFirst-- }
                else { $nAgeCrit-- }
                $diff++
            }
        }
    }

    Write-Host "  Allocation: Age-Critical=$nAgeCrit, First Flights=$nFirst, Progressing=$nProg" -ForegroundColor Gray

    # Exclude COMPLETED cadets from scheduling (they've already completed all 5 flights)
    $eligible = $prioritized | Where-Object { $_.Tier -ne 'COMPLETED' }

    $ageCritical = $eligible | Where-Object { $_.MonthsUntil18 -le 6 -and $_.FlightsCompleted -lt 5 }
    $firstFlights= $eligible | Where-Object { $_.FlightsCompleted -eq 0 -and ($_.CAPID -notin $ageCritical.CAPID) }
    $progressing = $eligible | Where-Object { $_.FlightsCompleted -ge 1 -and $_.FlightsCompleted -lt 5 -and ($_.CAPID -notin $ageCritical.CAPID) }

    $sched = [System.Collections.ArrayList]::new()
    $sqCounts = @{}

    function Add-WithCap {
        param(
            [Parameter(Mandatory=$false)][AllowNull()][System.Collections.IEnumerable]$List,
            [Parameter(Mandatory=$true)][int]$MaxCount,
            [Parameter(Mandatory=$true)][AllowEmptyCollection()][System.Collections.ArrayList]$Collector,
            [Hashtable]$SquadronCounts,
            [int]$Cap
        )
        # Handle null or empty list
        if ($null -eq $List) { return }

        foreach ($row in $List) {
            if ($Collector.Count -ge $MaxCount) { break }
            if ($Cap -and $SquadronCounts.ContainsKey($row.Squadron) -and $SquadronCounts[$row.Squadron] -ge $Cap) { continue }
            [void]$Collector.Add($row)
            if ($Cap) {
                if (-not $SquadronCounts.ContainsKey($row.Squadron)) { $SquadronCounts[$row.Squadron] = 0 }
                $SquadronCounts[$row.Squadron]++
            }
        }
    }

    # 1) Age-critical first
    Add-WithCap -List $ageCritical -MaxCount $nAgeCrit -Collector $sched -SquadronCounts $sqCounts -Cap $MaxPerSquadron

    # 2) First flights
    Add-WithCap -List $firstFlights -MaxCount ($nAgeCrit + $nFirst) -Collector $sched -SquadronCounts $sqCounts -Cap $MaxPerSquadron

    # 3) Progressing (2nd–4th)
    Add-WithCap -List $progressing -MaxCount $TotalSlots -Collector $sched -SquadronCounts $sqCounts -Cap $MaxPerSquadron

    # Top-up if underfilled (only from eligible cadets, not COMPLETED)
    if ($sched.Count -lt $TotalSlots) {
        $remaining = $eligible | Where-Object { $_.CAPID -notin ($sched | ForEach-Object CAPID) }
        Add-WithCap -List $remaining -MaxCount $TotalSlots -Collector $sched -SquadronCounts $sqCounts -Cap $MaxPerSquadron
    }

    # Sort schedule by squadron
    $schedBySquadron = $sched | Sort-Object Squadron, @{Expression='PriorityScore';Descending=$true}

    $ranked = $schedBySquadron | Select-Object `
        @{n='Slot';e={[array]::IndexOf($schedBySquadron, $_) + 1}},
        CAPID,Squadron,LastName,FirstName,FlightsCompleted,NextFlightNumber,
        PriorityScore,Tier,MonthsUntil18,DaysSinceLastFlight,DaysSinceJoin,LastFlightDate,JoinedDate

    $ranked | Export-Csv -NoTypeInformation -Path $OutputScheduleCsv -Encoding UTF8
    Write-Host "✅ Wrote schedule ($($ranked.Count) slots, sorted by squadron) to $OutputScheduleCsv" -ForegroundColor Green

    # Display schedule summary by squadron
    $schedSquadronSummary = $ranked | Group-Object Squadron | Sort-Object Name
    Write-Host "`nSchedule by Squadron:" -ForegroundColor Cyan
    foreach ($sq in $schedSquadronSummary) {
        Write-Host "  Squadron $($sq.Name): $($sq.Count) flights scheduled" -ForegroundColor White
    }
}

# ---------- Optional: Save to Cosmos DB ----------
if ($SaveToCosmosDb) {
    Write-Host "`nSaving results to Cosmos DB..." -ForegroundColor Cyan

    if (-not $ConnectionString -or -not $Database) {
        Write-Host "❌ Error: Cosmos DB configuration incomplete. Please set ConnectionString and Database parameters or environment variables." -ForegroundColor Red
    } else {
        # Prepare priority list data for Cosmos DB
        $calculatedDate = $AsOf.ToString('yyyy-MM-dd')

        # Group by squadron for squadron-level metrics
        $squadronMetrics = @{}
        foreach ($cadet in $prioritized) {
            $sq = $cadet.Squadron
            if (-not $squadronMetrics.ContainsKey($sq)) {
                $squadronMetrics[$sq] = @{
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

            $squadronMetrics[$sq].totalCadets++
            $squadronMetrics[$sq].byTier[$cadet.Tier]++
            $squadronMetrics[$sq].cadets += [PSCustomObject]@{
                capid = $cadet.CAPID
                lastName = $cadet.LastName
                firstName = $cadet.FirstName
                flightsCompleted = $cadet.FlightsCompleted
                nextFlightNumber = $cadet.NextFlightNumber
                priorityScore = $cadet.PriorityScore
                tier = $cadet.Tier
                monthsUntil18 = $cadet.MonthsUntil18
                daysSinceJoin = $cadet.DaysSinceJoin
                daysSinceLastFlight = $cadet.DaysSinceLastFlight
            }
        }

        # Calculate average priority scores
        foreach ($sq in $squadronMetrics.Keys) {
            $scores = $squadronMetrics[$sq].cadets | ForEach-Object { $_.priorityScore }
            $squadronMetrics[$sq].avgPriorityScore = if ($scores.Count -gt 0) {
                [math]::Round(($scores | Measure-Object -Average).Average, 2)
            } else { 0 }
        }

        # Create the main priority metrics document
        $priorityMetric = [PSCustomObject]@{
            id = "oflight-priority-$calculatedDate"
            metricType = "oflight-priority"
            calculatedDate = $calculatedDate
            totalCadets = $prioritized.Count
            byTier = @{
                Critical = ($prioritized | Where-Object { $_.Tier -eq 'Critical' }).Count
                High = ($prioritized | Where-Object { $_.Tier -eq 'High' }).Count
                Medium = ($prioritized | Where-Object { $_.Tier -eq 'Medium' }).Count
                Low = ($prioritized | Where-Object { $_.Tier -eq 'Low' }).Count
                COMPLETED = ($prioritized | Where-Object { $_.Tier -eq 'COMPLETED' }).Count
            }
            avgPriorityScore = [math]::Round(($prioritized | Measure-Object -Property PriorityScore -Average).Average, 2)
            squadrons = $squadronMetrics
            topPriority = ($prioritized | Select-Object -First 20 | ForEach-Object {
                [PSCustomObject]@{
                    capid = $_.CAPID
                    squadron = $_.Squadron
                    lastName = $_.LastName
                    firstName = $_.FirstName
                    priorityScore = $_.PriorityScore
                    tier = $_.Tier
                    flightsCompleted = $_.FlightsCompleted
                }
            })
            calculatedAt = (Get-Date -Format o)
        }

        # Save priority metrics
        $saved = Save-MetricsItem -Item $priorityMetric -ConnectionString $ConnectionString -Database $Database

        if (-not $saved) {
            Write-Host "❌ Failed to save priority metrics to Cosmos DB" -ForegroundColor Red
        }

        # If schedule was generated, save schedule metrics too
        if ($TotalSlots -and $OutputScheduleCsv -and (Test-Path $OutputScheduleCsv)) {
            $scheduleMetric = [PSCustomObject]@{
                id = "oflight-schedule-$calculatedDate"
                metricType = "oflight-schedule"
                calculatedDate = $calculatedDate
                totalSlots = $TotalSlots
                slotsAllocated = $ranked.Count
                allocationStrategy = @{
                    ageCritical = $AllocationAgeCritical
                    firstFlights = $AllocationFirst
                    progressing = $AllocationProgress
                }
                maxPerSquadron = $MaxPerSquadron
                squadrons = ($ranked | Group-Object Squadron | ForEach-Object {
                    [PSCustomObject]@{
                        squadron = $_.Name
                        slotsAllocated = $_.Count
                    }
                })
                schedule = ($ranked | ForEach-Object {
                    [PSCustomObject]@{
                        slot = $_.Slot
                        capid = $_.CAPID
                        squadron = $_.Squadron
                        lastName = $_.LastName
                        firstName = $_.FirstName
                        priorityScore = $_.PriorityScore
                        tier = $_.Tier
                        nextFlightNumber = $_.NextFlightNumber
                    }
                })
                calculatedAt = (Get-Date -Format o)
            }

            $savedSchedule = Save-MetricsItem -Item $scheduleMetric -ConnectionString $ConnectionString -Database $Database

            if (-not $savedSchedule) {
                Write-Host "❌ Failed to save schedule metrics to Cosmos DB" -ForegroundColor Red
            }
        }
    }
}

Write-Host "`n✅ OFlight Priority calculation complete!" -ForegroundColor Green
