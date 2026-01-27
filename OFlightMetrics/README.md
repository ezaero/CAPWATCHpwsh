# OFlightMetrics Function

## Overview
This Azure Function calculates Orientation Flight metrics by squadron for two time periods:
- **Previous Month**: All flights from the previous calendar month
- **Fiscal Year**: All flights since October 1st of the current fiscal year

The function runs automatically on the **1st of each month at 8:00 AM MST**.

## What It Does

1. **Queries O-Flight Data**: Retrieves all O-Flight records from the CosmosDB `syllabus` container
2. **Maps to Squadrons**: Uses Azure AD user data to map each CAPID to their squadron
3. **Calculates Flight Metrics**: Groups flights by squadron and calculates:
   - Total number of flights
   - Number of unique cadets who flew
   - Breakdown by syllabus (6, 7, 8, 9, 10)
4. **Calculates Additional Metrics**:
   - Zero O-Flights per Squadron (count and percentage)
   - Syllabus Completion Rates (per squadron and wing-wide)
   - Monthly Flight Trends (past 24 months)
   - Year-over-Year Comparison
   - Squadron Rankings by Engagement Rate
   - Time-to-First-Flight metrics
   - Squadron Leaderboards
   - Squadron Participation Details Report
   - **OFlight Priority Scores** (NEW)
5. **Saves to CosmosDB**: Writes all metrics to the `metrics` container for the web app to display

## Schedule

- **Timer Trigger**: `0 0 8 1 * *` (8:00 AM MST on the 1st of every month)
- On February 1st, it calculates January metrics
- On March 1st, it calculates February metrics, etc.

## Metrics Output Structure

### Previous Month Metric
```json
{
  "id": "monthly-2026-01",
  "metricType": "monthly-squadron",
  "period": "month",
  "year": 2026,
  "month": 1,
  "monthName": "January 2026",
  "startDate": "2026-01-01",
  "endDate": "2026-01-31",
  "totalFlights": 45,
  "totalUniqueCadets": 38,
  "squadrons": {
    "095": {
      "totalFlights": 12,
      "uniqueCadets": 10,
      "bySyllabus": {
        "6": 3,
        "7": 4,
        "8": 2,
        "9": 2,
        "10": 1
      }
    },
    "123": {
      "totalFlights": 8,
      "uniqueCadets": 7,
      "bySyllabus": {
        "6": 2,
        "7": 3,
        "8": 2,
        "9": 1,
        "10": 0
      }
    }
  },
  "calculatedAt": "2026-02-01T08:00:00Z"
}
```

### Fiscal Year Metric
```json
{
  "id": "fiscal-year-2026",
  "metricType": "fiscal-year-squadron",
  "period": "fiscal-year",
  "fiscalYear": 2026,
  "fiscalYearName": "FY2026",
  "startDate": "2025-10-01",
  "endDate": "2026-02-01",
  "totalFlights": 180,
  "totalUniqueCadets": 145,
  "squadrons": {
    "095": {
      "totalFlights": 45,
      "uniqueCadets": 38,
      "bySyllabus": {
        "6": 12,
        "7": 15,
        "8": 8,
        "9": 7,
        "10": 3
      }
    }
  },
  "calculatedAt": "2026-02-01T08:00:00Z"
}
```

## Environment Variables Required

- `CosmosDbConnectionString`: Connection string for Cosmos DB account
- `CosmosDbDatabase`: Database name (typically "orientation-flights")
- Azure Function Managed Identity must have Microsoft Graph permissions to read users

## Local Testing

Run the test script to execute the function locally:

```powershell
.\test-oflight-metrics.ps1
```

This will:
1. Verify Cosmos DB configuration
2. Authenticate to Microsoft Graph
3. Run the metrics calculation
4. Display results in the console
5. Write metrics to CosmosDB

## Deployment

This function is deployed as part of the CAPWATCH Azure Function App. After creating or modifying:

1. Commit changes to git
2. Push to Azure Function App
3. Verify the timer trigger is enabled in Azure Portal

## Querying Metrics from Web App

The COFLICS web app can query the metrics container:

### Get Latest Monthly Metrics
```sql
SELECT * FROM c
WHERE c.metricType = 'monthly-squadron'
ORDER BY c.year DESC, c.month DESC
OFFSET 0 LIMIT 1
```

### Get Current Fiscal Year Metrics
```sql
SELECT * FROM c
WHERE c.metricType = 'fiscal-year-squadron'
ORDER BY c.fiscalYear DESC
OFFSET 0 LIMIT 1
```

### Get Specific Month
```sql
SELECT * FROM c
WHERE c.id = 'monthly-2026-01'
```

## OFlight Priority Integration

As of the latest version, this function also calculates **OFlight Priority scores** for all active cadets. The priority algorithm considers:

- **First Flight Urgency (A)**: How long cadets have waited for their first flight (0-100 points)
- **Time Since Last Flight (B)**: Time elapsed since most recent flight (0-40 points)
- **Progression Equity (C)**: Priority for cadets with fewer completed flights (0-30 points)
- **Age Urgency (D)**: Critical priority for cadets approaching age 18 (0-40 points)

The priority calculation:
- Uses Azure AD `createdDateTime` as join date (DOB not available)
- Calculates flights from existing `allFlights` data
- Groups results by squadron
- Saves to `oflight-priority-YYYY-MM-DD` document in Cosmos DB

### Priority Metrics Output Structure

```json
{
  "id": "oflight-priority-2026-01-26",
  "metricType": "oflight-priority",
  "calculatedDate": "2026-01-26",
  "totalCadets": 245,
  "byTier": {
    "Critical": 12,
    "High": 34,
    "Medium": 156,
    "Low": 43
  },
  "avgPriorityScore": 42.5,
  "squadrons": {
    "095": {
      "totalCadets": 45,
      "byTier": {
        "Critical": 2,
        "High": 8,
        "Medium": 30,
        "Low": 5
      },
      "avgPriorityScore": 38.2,
      "cadets": [
        {
          "capid": "123456",
          "lastName": "Doe",
          "firstName": "John",
          "flightsCompleted": 0,
          "nextFlightNumber": 1,
          "priorityScore": 85.5,
          "tier": "Critical",
          "monthsUntil18": 5,
          "daysSinceJoin": 120,
          "daysSinceLastFlight": null
        }
      ]
    }
  },
  "topPriority": [
    {
      "capid": "123456",
      "squadron": "095",
      "lastName": "Doe",
      "firstName": "John",
      "priorityScore": 85.5,
      "tier": "Critical",
      "flightsCompleted": 0
    }
  ],
  "calculatedAt": "2026-01-26T15:30:00Z"
}
```

### Priority Tiers

- **Critical**: Zero flights >60 days, OR has flights but >180 days since last, OR <6 months until age 18 with <5 flights
- **High**: Zero flights 30-60 days, OR 90-180 days since last, OR 6-12 months until age 18 with <5 flights
- **Medium**: Has <5 flights completed (not Critical or High)
- **Low**: Has 5+ flights completed

## Notes

- Squadrons are extracted from user `companyName` field (format: "CO-095")
- Only flights with valid squadron mappings are included in metrics
- Metrics are upserted, so running the function multiple times overwrites previous data
- Fiscal Year runs from October 1 to September 30
- Priority scores are calculated automatically each time the function runs
- Priority calculation uses Azure AD join date (createdDateTime); DOB-based age urgency is limited without Member.txt data
