# DLDutyPos

Builds intentional Exchange distribution lists from CAPWATCH-derived Entra security groups.

The function reads a distribution-list map, expands source security-group members, applies exclusions, resolves user mail addresses, and queues Exchange distribution-list updates through the existing `DLGroupUpdateWorker`.

The source groups are not email lists. They are identity groups maintained by `CAPGroups`.

This function is triggered from the `dl-dutypos-plans` queue after `CAPGroups` completes successfully.

## Configuration

Create a JSON file based on `distribution-lists.example.json`, then point the function to it with:

```text
DLDUTYPOS_CONFIG_PATH=/path/to/distribution-lists.json
```

If no config file is present, no distribution-list jobs are queued.

Example:

```json
{
  "displayName": "CO Wing Staff",
  "email": "staff@cowg.cap.gov",
  "ensureGroup": false,
  "allowMissingSourceGroups": true,
  "includeGroupPatterns": [
    "CAP-Duty-WING-CO-*"
  ],
  "excludeGroups": []
}
```

If an included or excluded source group is missing, the distribution list is skipped unless the definition sets `allowMissingSourceGroups` to `true`.

`includeGroups` and `excludeGroups` match exact Entra security-group display names. `includeGroupPatterns` and `excludeGroupPatterns` use PowerShell wildcard matching after resolving candidate Entra group display names by prefix.

`ensureGroup` defaults to `false`, so DLDutyPos reuses existing Exchange distribution groups unless a definition explicitly opts in to creating a missing group.

