# OFlights Function

This function processes O-Flight data from CAPWATCH and syncs it to Azure Cosmos DB.

## Overview

The OFlights function:
1. Imports O-Flight records from the `OFlight.txt` file in the CAPWATCH data directory
2. Filters flights for specific syllabuses (6, 7, 8, 9, 10)
3. Groups flights by CAPID and Syllabus to identify first flight dates per syllabus
4. Validates that all CAPIDs exist in the current Azure AD user list
5. Syncs flight data to Cosmos DB with intelligent change detection
6. Only updates Cosmos DB documents when data has actually changed

## Configuration

Ensure the following environment variables are set:

- `CosmosDbConnectionString`: Connection string to your Cosmos DB account
- `CosmosDbDatabase`: Cosmos DB database name
- `CosmosDbContainer`: Cosmos DB container name for O-Flight records

## Cosmos DB Document Structure

Documents are stored with the following structure:

```json
{
  "id": "CAPID-Syllabus",
  "CAPID": "123456",
  "Syllabus": "6",
  "FirstFlight": "2023-07-09T00:00:00",
  "LastUpdated": "2024-01-24T10:30:00",
  "SyncSource": "OFlights"
}
```

## Change Detection Logic

The function implements a **compare-and-differential model**:
- **Query Phase**: Retrieves existing documents from Cosmos DB with their FirstFlight dates
- **Compare Phase**: Compares expected state (from OFlight.txt) with actual state (Cosmos DB)
- **Classify Phase**: Documents are classified into three groups:
  - **Orphaned**: Exist in Cosmos but removed from OFlight.txt → Delete
  - **Changed**: Exist in both but FirstFlight date differs → Upsert
  - **New**: Don't exist in Cosmos → Upsert
  - **Unchanged**: Exist in both with same FirstFlight → Skip
- **Sync Phase**: Only deletes orphaned docs, only upserts new/changed docs
- **Efficiency**: Skips unchanged documents, minimizing RU consumption

**Benefits:**
- ✅ No orphaned documents (auto-deleted)
- ✅ No wasted writes on unchanged data
- ✅ Lower Cosmos DB RU costs
- ✅ Cost-effective for frequent (daily) runs
- ✅ Smart change detection prevents unnecessary upserts

## Schedule

By default, this function runs daily at 2:00 AM UTC (via the timer trigger in `function.json`).

To change the schedule, modify the `schedule` property in `function.json` using [CRON expressions](https://learn.microsoft.com/en-us/azure/azure-functions/functions-bindings-timer?tabs=in-process%2Cv1%2Cbroadcast&pivots=programming-language-powershell#ncrontab-expressions).

## Logging

All operations are logged to `$env:HOME/logs/script_log_YYYY-MM-DD.txt` with timestamps and detailed status messages.

## Requirements

- CAPWATCH data directory must be accessible at `$HOME/data/CAPWatch`
- `OFlight.txt` file must be present in the CAPWATCH data directory
- Azure AD/Microsoft Graph connection must be established
- Cosmos DB account must exist and be accessible
