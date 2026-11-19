# CAPGroups

Maintains CAPWATCH-derived Microsoft Entra security groups as a durable identity layer.

This function is triggered from the `capwatch-post-download` queue after `download-extract-capwatch` completes successfully.

`CAPGroups` creates source-of-truth identity groups. These groups are not email lists. Downstream functions, such as `DLDutyPos`, can use them to build intentional Exchange distribution lists only where there is a communications need.

## Source Data

The function reads:

- `Member.txt`
- `DutyPosition.txt`
- `Achievements.txt`
- `MbrAchievements.txt`
- `MbrTasks.txt`
- `Organization.txt`

CAPIDs are matched to Entra users through `employeeId`. Unit-scoped and group-scoped plans use Entra profile attributes such as `companyName`, `employeeType`, and `department` where needed.

## Group Categories

The function creates and reconciles non-mail-enabled Entra security groups such as:

- `CAP-Wing-CO`
- `CAP-Region-RMR`
- `CAP-Unit-CO-162`
- `CAP-MemberType-SENIOR`
- `CAP-MemberStatus-ACTIVE`
- `CAP-WingAudience-CO-Cadets`
- `CAP-GroupAudience-CO-4-All`
- `CAP-UnitAudience-CO-162-Cadets`
- `CAP-Duty-WING-CO-IT-Director-of-IT`
- `CAP-OpsQual-GES-General-Emergency-Services`
- `CAP-OpsQual-Mission-Pilot`
- `CAP-OpsQual-ES-Group-4`

Each plan includes the desired member IDs. During execution, `CAPGroups` creates missing groups, adds desired members, and removes stale members.

## Ops Qual Groups

`CAPGroups` now creates exact per-achievement Ops Qual groups for the configured AchvIDs in `CAPGroups.Helpers.ps1`. The group display name comes from `Achievements.txt`, normalized into the `CAP-OpsQual-*` naming convention. Numeric AchvIDs are intentionally omitted from the group display name because they are not useful to end users.

Examples:

- `CAP-OpsQual-VFR-Pilot`
- `CAP-OpsQual-GES-General-Emergency-Services`
- `CAP-OpsQual-MS-Mission-Scanner`
- `CAP-OpsQual-MP-SAR-DR-Mission-Pilot`
- `CAP-OpsQual-ICUT-Introductory-Communications-User-Training`

Exact per-achievement groups include members whose `MbrAchievements.txt` row for that AchvID has `Status` of `ACTIVE` or `TRAINING`.

A `SUSPENDED` row takes precedence over other rows for the same CAPID and achievement or task. A member can still qualify for a combined group through another eligible qualification. Manual KAPA inclusions also honor explicit suspensions of the pilot task.

The function also keeps curated Ops Qual rollup groups for operational use cases that are not a one-to-one achievement mapping:

- `CAP-OpsQual-Pilot`: active pilot task, `MbrTasks` `TaskID 69`
- `CAP-OpsQual-KAPA-Pilot`: active pilots limited to configured KAPA units plus approved manual CAPIDs
- `CAP-OpsQual-ES`: active General Emergency Services
- `CAP-OpsQual-ES-Group-4`: active General Emergency Services limited to active Group 4 units
- `CAP-OpsQual-Mission-Pilot`: mission pilot qualification intersected with active pilot task
- `CAP-OpsQual-Mission-Pilot-MFC`: active mission pilots with active Mountain Flying Certification
- `CAP-OpsQual-Mission-Base-Staff`: combined mission base staff qualifications
- `CAP-OpsQual-sUAS`: combined sUAS qualifications

If an exact per-achievement group and a curated rollup group would produce the same display name, the exact achievement group is kept and the duplicate plan is skipped.

## Execution

By default the function runs in preview mode. Set this app setting to apply changes:

```text
CAP_GROUPS_EXECUTE=true
```

Preview mode logs planned group creation and membership changes without writing to Graph.

Large plan sets are processed in batches. Set `CAP_GROUPS_PLAN_BATCH_SIZE` to override the default batch size of `35`. When more plans remain, the function queues a continuation message back to `capwatch-post-download` with the next plan index.

`DLDutyPos` is queued only after this function completes. In preview mode, the downstream queue message is suppressed by default so distribution lists are not synced from stale or missing source groups. Set this only for downstream planner testing:

```text
CAP_GROUPS_QUEUE_DLDUTYPOS_IN_PREVIEW=true
```

## Permissions

The managed identity needs Microsoft Graph permissions that allow it to read users and create/update group membership, including:

- `User.Read.All`
- `Group.ReadWrite.All`
- `GroupMember.ReadWrite.All`

## Validation

Run the focused test before deployment:

```powershell
pwsh -NoProfile -File CAPGroups/CAPGroups.Tests.ps1
```

