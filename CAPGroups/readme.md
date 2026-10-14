# CAPGroups

Maintains CAPWATCH-derived Entra security groups as a durable identity layer.

This function is triggered from the `capwatch-post-download` queue after `download-extract-capwatch` completes successfully.

The function reads:

- `Member.txt`
- `DutyPosition.txt`
- `Organization.txt`

It creates and reconciles non-mail-enabled Entra security groups such as:

- `CAP-Wing-CO`
- `CAP-Region-RMR`
- `CAP-Unit-CO-162`
- `CAP-MemberType-SENIOR`
- `CAP-MemberStatus-ACTIVE`
- `CAP-Duty-WING-CO-IT-Director-of-IT`

These groups are intended to be source-of-truth identity groups. They are not intended to be email lists.

## Execution

By default the function runs in preview mode. Set this app setting to apply changes:

```text
CAP_GROUPS_EXECUTE=true
```

Preview mode logs planned group creation and membership changes without writing to Graph.

`DLDutyPos` is queued only after this function completes. In preview mode, the downstream queue message is suppressed by default so distribution lists are not synced from stale or missing source groups. Set this only for downstream planner testing:

```text
CAP_GROUPS_QUEUE_DLDUTYPOS_IN_PREVIEW=true
```

