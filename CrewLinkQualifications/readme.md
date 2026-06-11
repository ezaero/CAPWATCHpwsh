# CrewLinkQualifications

Timer-triggered Azure Function that prepares CAPWATCH aircrew qualification snapshots for CREWLINK.

## Schedule

Runs on the same schedule as `OFlights`:

```text
0 0 14 * * 1,4
```

That is 1400 UTC every Monday and Thursday.

## Source Files

The function reads these files from `$HOME\data\CAPWatch`:

- `Achievements.txt` for the `AchvID` to achievement name lookup.
- `MbrAchievements.txt` for member qualification status records.

## Status Mapping

- `ACTIVE` becomes `FULLY_QUALIFIED` and can satisfy proficiency mission matching.
- `TRAINING` remains `TRAINING` and can satisfy training opportunity visibility.
- All other statuses become `NOT_ELIGIBLE` and should not satisfy automated matching.

The snapshot retains non-eligible records for visibility and diagnostics.

## Output

Documents are upserted to Cosmos DB with:

- `id`: `CAPID-AchvID`
- `CAPID`: partition key value
- `SyncSource`: `CrewLinkQualifications`
- `EligibilityStatus`: `FULLY_QUALIFIED`, `TRAINING`, or `NOT_ELIGIBLE`

By default the function writes to the `crewlinkQualifications` container. Override with the `CrewLinkQualificationsCosmosDbContainer` app setting when needed. The container should use `/CAPID` as its partition key.

## Sync Behavior

The function treats CAPWATCH as the source of truth:

- Builds a deduplicated desired snapshot keyed by `CAPID-AchvID`.
- Queries existing `CrewLinkQualifications` documents from Cosmos DB.
- Upserts only new or meaningfully changed documents.
- Deletes stale documents that are no longer present in the CAPWATCH snapshot.

Timestamp fields such as `LastUpdated` and `timestamp` are intentionally ignored during comparison so unchanged qualifications do not get rewritten on every run.
