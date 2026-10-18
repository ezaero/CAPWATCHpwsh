# Distribution List Configuration

`DLDutyPos` reads every `*.yaml` file in this folder and treats the enabled definitions as one configuration set. Enabled definitions must have unique `displayName` and `email` values.

The source groups named in these files are Entra security groups maintained from CAPWATCH. The Exchange distribution lists are the communication layer; they should exist only where there is a real communications need.

## Files

- `duty-positions.yaml` is active today. It builds staff and duty-position lists from `CAP-Duty-*` source groups.
- `units.yaml` is populated from the current Colorado Wing CAPWATCH `Organization.txt`, but entries are disabled until unit lists move away from `DLSeniorsCadets`. It uses `CAP-*Audience-*` source groups so cadet lists can include cadets, parents, and CP/EX staff.
- `ops-quals.yaml` documents the current `DLOpsQuals` lists and criteria, but entries are disabled until matching `CAP-OpsQual-*` source groups are created by `CAPGroups`.
- `specialty-tracks.yaml`, `manual.yaml`, and `announcements.yaml` are placeholders for later categories.

## Adding a Duty-Position List

1. Add the definition to `duty-positions.yaml`.
2. Use existing `CAP-Duty-*` source groups in `includeGroups`, or use `includeGroupPatterns` for a controlled wildcard.
3. Set `email` to the existing distribution list SMTP address if the list already exists.
4. Leave `ensureGroup: false` when reusing an existing list. Set it to `true` only when the function should create a missing distribution list.
5. Run `DLDutyPos/DLDutyPos.Tests.ps1` before deployment.

## Adding an Ops-Qual List

1. Add or update the source-group creation in `CAPGroups` first, so the desired membership exists as an Entra security group.
2. Add the distribution list definition to `ops-quals.yaml`.
3. Record the CAPWATCH rule in `criteria` so future maintainers can compare the YAML to the source-group logic.
4. Set `enabled: true` only after the source group exists and the old `DLOpsQuals` updater no longer manages the same distribution list.
5. Run a preview and compare expected member counts before the first live run.

## Adding a Unit, Group, or Wing List

1. Add the definition to `units.yaml`.
2. For a wing list, include a `CAP-WingAudience-CO-*` source group.
3. For a group list, include a `CAP-GroupAudience-CO-#-All` source group.
4. For a unit list, include exactly one `CAP-UnitAudience-CO-###-*` source group.
5. Use the `Cadets` audience when the current list should include cadets, parents, and CP/EX staff.
6. Enable the list only after `DLSeniorsCadets` stops managing that same distribution list.

## Removing a List

1. Remove the definition from the YAML file, or set `enabled: false` if you want to keep a migration note.
2. Confirm no other function still manages the same distribution list.
3. Removing a definition stops future reconciler updates; it does not delete the Exchange distribution list.
4. Retire or delete the Exchange list separately if it is no longer needed.
