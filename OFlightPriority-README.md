# CAP Orientation Flight Priority System

## Overview

The OFlight Priority System calculates priority scores for cadets to determine who should receive orientation flights. The system considers multiple factors including:

- **First Flight Urgency (A)**: How long a cadet has waited for their first flight
- **Time Since Last Flight (B)**: How long since their most recent flight
- **Progression Equity (C)**: Priority to cadets with fewer completed flights
- **Age Urgency (D)**: Priority for cadets approaching age 18

## Files

- **Get-OFlightPriority.ps1**: Main prioritization script
- **test-oflight-priority.ps1**: Interactive test script for local execution
- **OFlightPriority.csv**: Output file with full prioritized list
- **OFlightSchedule.csv**: Optional output file with flight schedule

## Data Sources

The script reads from two CAPWATCH data files:

1. **Member.txt** (`$env:HOME\data\CAPWatch\Member.txt`)
   - Fields used: CAPID, NameFirst, NameLast, Unit (Squadron), Joined, DOB, MbrStatus, Type
   - Filters: Active cadets only

2. **OFlight.txt** (`$env:HOME\data\CAPWatch\OFlight.txt`)
   - Fields used: CAPID, Syllabus, FltDate
   - Filters: Syllabus 6-10 only (orientation flight syllabi)

## Priority Scoring

### Components

1. **A - First Flight Urgency (0-100 points)**
   - Only applies to cadets with zero flights
   - Score increases proportionally with days since join date
   - Target: First flight within 60 days of joining
   - Formula: `(DaysSinceJoin / 60) * 100` (capped at 100)

2. **B - Time Since Last Flight (0-40 points)**
   - Only applies to cadets with 1+ flights
   - Approximately 10 points per month since last flight
   - Formula: `Min(40, (DaysSinceLast / 30) * 10)`

3. **C - Progression Equity (0-30 points)**
   - Rewards cadets with fewer completed flights
   - Formula: `(5 - FlightsCompleted) * 6`
   - 0 flights = 30 points, 1 flight = 24 points, etc.

4. **D - Age Urgency (0-40 points)**
   - Critical priority for cadets approaching age 18
   - 0-3 months until 18th birthday: 40 points
   - 3-6 months: 30 points
   - 6-12 months: 20 points
   - 12-18 months: 10 points
   - 18+ months: 0 points

### Total Score

`PriorityScore = A + B + C + D` (Range: 0-210)

### Tie-Breaking

When cadets have the same priority score, ties are broken by:
1. Fewest flights completed (higher priority)
2. Most days since last flight (higher priority)
3. Age in years (older gets priority)
4. Deterministic jitter (consistent hash of CAPID+Name)

### Priority Tiers

Cadets are classified into tiers for easy identification:

- **Critical**:
  - Zero flights AND >60 days since join, OR
  - Has flights AND 180+ days since last flight, OR
  - <6 months until age 18 AND <5 flights completed

- **High**:
  - Zero flights AND 30-60 days since join, OR
  - Has flights AND 90-180 days since last flight, OR
  - 6-12 months until age 18 AND <5 flights completed

- **Medium**:
  - Has <5 flights completed (not Critical or High)

- **Low**:
  - Has 5+ flights completed

## Usage

### Basic Usage (Priority List Only)

```powershell
.\Get-OFlightPriority.ps1 -OutputCsv ".\OFlightPriority.csv"
```

### With Cosmos DB Storage

```powershell
.\Get-OFlightPriority.ps1 `
    -OutputCsv ".\OFlightPriority.csv" `
    -SaveToCosmosDb
```

### With Flight Schedule

```powershell
.\Get-OFlightPriority.ps1 `
    -OutputCsv ".\OFlightPriority.csv" `
    -OutputScheduleCsv ".\OFlightSchedule.csv" `
    -TotalSlots 20
```

### With Squadron Cap

```powershell
.\Get-OFlightPriority.ps1 `
    -OutputCsv ".\OFlightPriority.csv" `
    -OutputScheduleCsv ".\OFlightSchedule.csv" `
    -TotalSlots 20 `
    -MaxPerSquadron 3
```

### Full Example with Cosmos DB

```powershell
.\Get-OFlightPriority.ps1 `
    -OutputCsv ".\OFlightPriority.csv" `
    -OutputScheduleCsv ".\OFlightSchedule.csv" `
    -TotalSlots 20 `
    -MaxPerSquadron 3 `
    -SaveToCosmosDb
```

### Interactive Testing

```powershell
.\test-oflight-priority.ps1
```

## Parameters

### Required
- **OutputCsv**: Path for the full prioritized cadet list (CSV)

### Optional
- **MemberPath**: Path to Member.txt (default: `$env:HOME\data\CAPWatch\Member.txt`)
- **OFlightsPath**: Path to OFlight.txt (default: `$env:HOME\data\CAPWatch\OFlight.txt`)
- **OutputScheduleCsv**: Path for flight schedule CSV (optional)
- **AsOf**: Reference date for calculations (default: today)
- **TotalSlots**: Number of flight slots to schedule (required for schedule generation)
- **MaxPerSquadron**: Maximum flights per squadron in schedule (optional)
- **FirstFlightDaysThreshold**: Target days for first flight (default: 60)

### Allocation Parameters (for Schedule)
- **AllocationFirst**: Fraction for first flights (default: 0.40)
- **AllocationProgress**: Fraction for progressing flights (default: 0.40)
- **AllocationAgeCritical**: Fraction for age-critical cadets (default: 0.20)

### Cosmos DB Parameters
- **SaveToCosmosDb**: Switch to enable saving results to Cosmos DB Metrics container
- **ConnectionString**: Cosmos DB connection string (default: `$env:CosmosDbConnectionString`)
- **Database**: Cosmos DB database name (default: `$env:CosmosDbDatabase`)

## Output Files

### OFlightPriority.csv

Full prioritized list of all active cadets, sorted by squadron then priority score.

**Columns:**
- CAPID
- Squadron
- LastName, FirstName
- FlightsCompleted (0-5)
- NextFlightNumber (1-5)
- JoinedDate, DaysSinceJoin
- LastFlightDate, DaysSinceLastFlight
- DOB, AgeYears, MonthsUntil18
- A_FirstFlightUrgency, B_SinceLastFlight, C_ProgressionEquity, D_AgeUrgency
- PriorityScore (total)
- Tier (Critical/High/Medium/Low)

### OFlightSchedule.csv (Optional)

Scheduled flight list based on slot allocation strategy.

**Columns:**
- Slot (1-N)
- CAPID, Squadron
- LastName, FirstName
- FlightsCompleted, NextFlightNumber
- PriorityScore, Tier
- MonthsUntil18
- DaysSinceLastFlight, DaysSinceJoin
- LastFlightDate, JoinedDate

## Schedule Allocation Strategy

When creating a flight schedule, slots are allocated in three categories:

1. **Age-Critical (20% by default)**
   - Cadets with <6 months until 18th birthday AND <5 flights

2. **First Flights (40% by default)**
   - Cadets with zero flights (excluding age-critical)

3. **Progressing (40% by default)**
   - Cadets with 1-4 flights (excluding age-critical)

Within each category, cadets are selected by priority score. If a squadron cap is set, no squadron will receive more than the specified number of slots.

## Cosmos DB Storage

When `-SaveToCosmosDb` is enabled, the script saves two types of documents to the Cosmos DB Metrics container:

### 1. Priority Metrics Document

**Document ID**: `oflight-priority-YYYY-MM-DD`
**Metric Type**: `oflight-priority` (partition key)

**Contents**:
- **totalCadets**: Total number of active cadets
- **byTier**: Count of cadets in each tier (Critical, High, Medium, Low)
- **avgPriorityScore**: Average priority score across all cadets
- **squadrons**: Per-squadron breakdown with:
  - Total cadets per squadron
  - Tier distribution per squadron
  - Average priority score per squadron
  - Full list of cadets with priority details
- **topPriority**: Top 20 highest-priority cadets wing-wide
- **calculatedAt**: ISO 8601 timestamp of calculation

### 2. Schedule Metrics Document (if schedule generated)

**Document ID**: `oflight-schedule-YYYY-MM-DD`
**Metric Type**: `oflight-schedule` (partition key)

**Contents**:
- **totalSlots**: Total slots requested
- **slotsAllocated**: Actual slots filled
- **allocationStrategy**: Allocation percentages used
- **maxPerSquadron**: Squadron cap (if set)
- **squadrons**: Slots allocated per squadron
- **schedule**: Complete flight schedule with:
  - Slot number
  - CAPID, name, squadron
  - Priority score, tier
  - Next flight number
- **calculatedAt**: ISO 8601 timestamp

### Querying from Cosmos DB

These documents can be queried for dashboards or reports:

```sql
-- Get latest priority metrics
SELECT * FROM c
WHERE c.metricType = 'oflight-priority'
ORDER BY c.calculatedDate DESC

-- Get all schedules
SELECT * FROM c
WHERE c.metricType = 'oflight-schedule'
ORDER BY c.calculatedDate DESC

-- Get squadron-specific priority data
SELECT c.squadrons['123'] FROM c
WHERE c.metricType = 'oflight-priority'
AND c.calculatedDate = '2026-01-26'
```

## Example Scenarios

### Scenario 1: New Cadet (Just Joined)
- Flights: 0
- Days Since Join: 10
- Months Until 18: 36

**Score:**
- A (First Flight): 16.7 points (10/60 * 100)
- B (Since Last): 0 points (no previous flights)
- C (Progression): 30 points (5-0 * 6)
- D (Age Urgency): 0 points (>18 months)
- **Total: 46.7 points** → Tier: Medium

### Scenario 2: Cadet Approaching Age 18
- Flights: 2
- Days Since Last Flight: 45
- Months Until 18: 5

**Score:**
- A (First Flight): 0 points (has flights)
- B (Since Last): 15 points (45/30 * 10)
- C (Progression): 18 points (5-2 * 6)
- D (Age Urgency): 30 points (3-6 months)
- **Total: 63 points** → Tier: Critical

### Scenario 3: Cadet with Long Wait
- Flights: 1
- Days Since Last Flight: 200
- Months Until 18: 24

**Score:**
- A (First Flight): 0 points (has flights)
- B (Since Last): 40 points (capped at max)
- C (Progression): 24 points (5-1 * 6)
- D (Age Urgency): 0 points (>18 months)
- **Total: 64 points** → Tier: Critical

## Notes

- The system automatically filters for active cadets only
- Flights beyond 5 are not counted (5-for-5 completion)
- Only orientation flight syllabi (6-10) are considered
- Output is sorted by squadron for easy squadron commander review
- Jitter ensures consistent but random-appearing tie-breaking

## Integration with Azure Functions

This script is designed to work with the existing CAPWATCH data pipeline:
- Reads from the same data directory as checkAccounts and OFlights functions
- Can be scheduled as a separate Azure Function if needed
- Compatible with Cosmos DB storage for results

## Author

Michael Schulte, Capt
Colorado Wing, Civil Air Patrol
