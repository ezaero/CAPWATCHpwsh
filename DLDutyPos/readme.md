# DLDutyPos

Builds intentional Exchange distribution lists from CAPWATCH-derived Entra security groups.

The function reads a distribution-list map, expands source security-group members, applies exclusions, resolves user mail addresses, and queues Exchange distribution-list updates through the existing `DLGroupUpdateWorker`.

The source groups are not email lists. They are identity groups maintained by `CAPGroups`.

This function is triggered from the `dl-dutypos-plans` queue after `CAPGroups` completes successfully.

## Configuration

By default, the function reads every `*.yaml` file in `DLDutyPos/distribution-lists/`, sorted by file name. You can also point the function at a single YAML or JSON file with:

```text
DLDUTYPOS_CONFIG_PATH=/path/to/distribution-lists
```

If no config path is present, no distribution-list jobs are queued.

Example:

```yaml
distributionLists:
  - displayName: "CO Wing Staff"
    email: "staff@cowg.cap.gov"
    ensureGroup: false
    allowMissingSourceGroups: true
    includeGroupPatterns:
      - "CAP-Duty-WING-CO-*"
    excludeGroups: []
```

If an included or excluded source group is missing, the distribution list is skipped unless the definition sets `allowMissingSourceGroups` to `true`.

`includeGroups` and `excludeGroups` match exact Entra security-group display names. `includeGroupPatterns` and `excludeGroupPatterns` use PowerShell wildcard matching after resolving candidate Entra group display names by prefix.

`enabled` defaults to `true`. Set `enabled: false` for migration notes or planned lists that should not queue updates yet.

`ensureGroup` defaults to `false`, so DLDutyPos reuses existing Exchange distribution groups unless a definition explicitly opts in to creating a missing group.

Enabled definitions must have unique `displayName` and `email` values across all YAML files in the folder.

See `distribution-lists/README.md` for the category layout and the process for adding or removing duty-position, unit, and ops-qual lists.

