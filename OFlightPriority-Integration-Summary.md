# OFlight Priority Integration Summary

## Overview

The OFlight Priority calculation has been successfully integrated into the **OFlightMetrics** Azure Function. Priority scores are now calculated automatically on the 1st of each month at 8:00 AM MST, along with all other orientation flight metrics.

## What Changed

### 1. OFlightMetrics Function (`OFlightMetrics/run.ps1`)
- **Added Priority Calculation**: The function now calculates priority scores for all active cadets after computing flight metrics
- **Uses Existing Data**: Leverages already-loaded flight data and user information from Azure AD
- **Automatic Execution**: Runs on the same schedule as flight metrics (monthly timer trigger)
- **Cosmos DB Storage**: Saves priority results to the `metrics` container as `oflight-priority-YYYY-MM-DD` documents

### 2. Standalone Script (`Get-OFlightPriority.ps1`)
- **Still Available**: Can be run manually for ad-hoc priority calculations
- **Supports Scheduling**: Useful for generating flight schedules with slot allocation
- **CSV Export**: Exports results to CSV files for offline review
- **Cosmos DB Optional**: Can optionally save to Cosmos DB with `-SaveToCosmosDb` flag

## Priority Calculation Details

### Components (0-210 points total)

- **A - First Flight Urgency (0-100)**: Priority for cadets waiting for their first flight
  - Based on days since joining CAP
  - Target: First flight within 60 days
  - Formula: `(DaysSinceJoin / 60) * 100` (capped at 100)

- **B - Time Since Last Flight (0-40)**: Time elapsed since most recent flight
  - ~10 points per month since last flight
  - Formula: `Min(40, (DaysSinceLast / 30) * 10)`

- **C - Progression Equity (0-30)**: Priority for cadets with fewer flights
  - Formula: `(5 - FlightsCompleted) * 6`

- **D - Age Urgency (0-40)**: Critical priority for cadets approaching age 18
  - 0-3 months: 40 points
  - 3-6 months: 30 points
  - 6-12 months: 20 points
  - 12-18 months: 10 points
  - 18+ months: 0 points

### Priority Tiers

- **Critical**: Urgent cases requiring immediate attention
- **High**: Should be scheduled soon
- **Medium**: Normal priority
- **Low**: Completed 5-for-5 program

## Data Sources

### Integrated (OFlightMetrics)
- **Flight Data**: From Cosmos DB `syllabus` container
- **User Data**: From Azure AD (Microsoft Graph API)
- **Join Date**: Uses `createdDateTime` from Azure AD
- **DOB**: From Azure AD `onPremisesExtensionAttributes/extensionAttribute1` (populated by checkAccounts function from CAPWATCH Member.txt)

### Standalone (Get-OFlightPriority.ps1)
- **Flight Data**: From `OFlight.txt` CAPWATCH file
- **User Data**: From `Member.txt` CAPWATCH file
- **Join Date**: From `Joined` field in Member.txt
- **DOB**: From `DOB` field in Member.txt (full age urgency)

## Cosmos DB Documents

### Priority Metrics Document
- **ID**: `oflight-priority-YYYY-MM-DD`
- **Partition Key**: `oflight-priority` (metricType)
- **Contents**:
  - Total cadets and average priority score
  - Tier breakdown (Critical, High, Medium, Low)
  - Per-squadron metrics with full cadet list
  - Top 20 highest-priority cadets wing-wide

### Schedule Metrics Document (Standalone only)
- **ID**: `oflight-schedule-YYYY-MM-DD`
- **Partition Key**: `oflight-schedule` (metricType)
- **Contents**:
  - Allocation strategy and slots
  - Per-squadron slot distribution
  - Complete flight schedule

## Usage

### Automatic Calculation (Recommended)
The priority scores are calculated automatically on the 1st of each month. No action required.

Query the latest priority data from Cosmos DB:
```sql
SELECT * FROM c
WHERE c.metricType = 'oflight-priority'
ORDER BY c.calculatedDate DESC
OFFSET 0 LIMIT 1
```

### Manual Calculation
For ad-hoc priority calculations or flight scheduling:

```powershell
# Priority list only
.\Get-OFlightPriority.ps1 -OutputCsv ".\OFlightPriority.csv"

# With Cosmos DB storage
.\Get-OFlightPriority.ps1 -OutputCsv ".\OFlightPriority.csv" -SaveToCosmosDb

# With flight schedule
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

## Benefits

1. **Automated Priority Tracking**: Priority scores update automatically each month
2. **Integrated with Metrics**: All OFlight data in one place
3. **Squadron-Level Insights**: See priority distribution per squadron
4. **Flexible Scheduling**: Standalone script supports flight scheduling with slot allocation
5. **Dashboard Ready**: Data structured for web dashboard consumption

## Limitations

### OFlightMetrics Integration
- **Join Date Source**: Uses Azure AD `createdDateTime` instead of CAP join date from Member.txt
  - May differ from actual CAP membership start date
- **DOB Accuracy**: Depends on checkAccounts function updating extensionAttribute1
  - If checkAccounts hasn't run recently, DOB may be stale

### Workaround
Use the standalone `Get-OFlightPriority.ps1` script with Member.txt for historical priority calculations when needed.

## Future Enhancements

1. **Email Notifications**: Send priority alerts for Critical tier cadets
2. **Auto-Schedule Generation**: Generate suggested flight schedules automatically
3. **Historical Tracking**: Track priority score changes over time
4. **Cache Member Data**: Implement caching mechanism to improve performance for large datasets

## Files Modified

- ✅ `OFlightMetrics/run.ps1` - Added priority calculation logic
- ✅ `OFlightMetrics/README.md` - Documented priority integration
- ✅ `Get-OFlightPriority.ps1` - Standalone priority script
- ✅ `test-oflight-priority.ps1` - Interactive test script
- ✅ `OFlightPriority-README.md` - Complete priority documentation

## Testing

Run the OFlightMetrics test to verify priority integration:
```powershell
.\test-oflight-metrics.ps1
```

Check for the priority calculation log entries and verify the Cosmos DB document is created.

---

**Author**: Michael Schulte, Capt
**Date**: January 2026
**Version**: 1.0
