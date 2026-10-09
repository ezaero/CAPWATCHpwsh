# Update Teams Script

## Overview

The `updateTeams` PowerShell script is designed to manage and synchronize Microsoft Teams for the Colorado Wing CAP. It ensures that Teams are created, updated, and populated with the correct members and owners based on CAPWATCH data. This script integrates with Microsoft Graph API and Azure services to automate the management of Teams.

---

## Features

1. **Team Management**:
   - Creates new Teams for units if they do not already exist.
   - Updates existing Teams with the correct members and owners.
   - Ensures aliases are generated and assigned to Teams.

2. **Member Synchronization**:
   - Compares unit members in CAPWATCH with current Team members.
   - Adds missing members to Teams.
   - Removes members who no longer belong to the unit.
   - Adds one-off external users from `ManualGuests.csv` to explicitly named Teams.
   - Adds cadets with CAC appointments from `CadetDutyPositions.txt` to the CAC Team.

3. **Owner Management**:
   - Ensures the unit commander is the owner of the Team.
   - Ensures senior members assigned as `Cadet Advisory Council Senior Advisor` in `DutyPosition.txt` are owners of the CAC Team.
   - Adds a default owner (e.g., Mike Schulte) to all Teams for script execution purposes.

4. **Azure Integration**:
   - Uses Microsoft Graph API for managing Teams, members, and owners.
   - Authenticates securely using Azure Managed Identity.

5. **Logging**:
   - Logs all actions, including Team creation, member updates, and errors.

---

## Prerequisites

- **Microsoft Graph PowerShell SDK**:
  - Ensure the SDK is installed and authenticated before running the script.

- **Azure Managed Identity**:
  - The Azure Function App must have permissions to manage Teams via Microsoft Graph API.

- **CAPWATCH Data**:
  - Ensure the CAPWATCH CSV files (`Organization.txt`, `Commanders.txt`, `DutyPosition.txt`, etc.) are available in the `$($env:HOME)\data\CAPWatch` directory.

---

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/your-repo/CAPWATCHSyncPWSH.git
   cd CAPWATCHSyncPWSH/updateTeams

## Key Functions

### Compare-UserIds
- Compares two arrays of user IDs and identifies:
  - Users present in both arrays.
  - Users only in the first array (to be added to Teams).
  - Users only in the second array (to be removed from Teams).

### GetAllUsers
- Retrieves all users from Microsoft Entra ID (Azure AD).

### GetAllGroups
- Retrieves all Microsoft 365 Groups that are associated with Teams.

### GetUnits
- Retrieves a list of all unit charter numbers and names in the Wing.

### GetCommander
- Retrieves the commander of a specific unit based on CAPWATCH data.

### CheckTeamExists
- Checks if a Team exists for a given unit.

### New-TeamAlias
- Generates a camel-cased alias for a Team based on the unit name.

### CheckTeams
- Ensures all required Teams exist and are properly configured.

### PopulateTeams
- Synchronizes members of Teams with CAPWATCH data.

---

## Logic Flow

### Retrieve Data
- Fetches unit data, users, and groups from CAPWATCH and Microsoft Graph API.

### Check Teams
- Ensures all required Teams exist and are properly configured.

### Populate Teams
- Compares unit members with Team members.
- Adds missing members to Teams.
- Removes members who no longer belong to the unit.

### Owner Management
- Ensures the unit commander is the owner of the Team.
- Adds a default owner to all Teams.

### Logging
- Logs all actions, including Team creation, member updates, and errors.

---

## Outputs

- **Logs**: All actions and errors are logged to the `$($env:HOME)\logs` directory.

---

## Function App Settings

- `EXECUTE=true`: applies membership changes. If omitted, the function runs in dry-run mode.
- `FORCE=true`: skips interactive prompts in non-interactive Azure Functions runs.
- `CREATE_MISSING_TEAMS=true`: creates Teams for active CAPWATCH units/groups that do not already have a Team whose display name starts with `CO-###`.
- `CREATE_MISSING_TEAM_UNITS=200`: optional comma-, semicolon-, or space-separated unit filter for Team creation. Use this to create only selected missing Teams, such as `200` or `CO-200`.
- `TEAM_SYNC_EXCLUDED_GROUP_IDS=c67c8903-14ec-4e31-a40d-d6d157c5b06b`: optional comma-, semicolon-, or space-separated Team group IDs to exclude from CAPWATCH unit-prefix membership sync. Use this for staff or special-purpose Teams whose display name starts with a unit prefix but should not receive all unit members.
- `CAC_TEAM_DISPLAY_NAME=CAC`: optional override for the Team display name used by the CAC duty-position sync. Defaults to `CAC`.

Current one-off exclusion: set `TEAM_SYNC_EXCLUDED_GROUP_IDS=c67c8903-14ec-4e31-a40d-d6d157c5b06b` to keep `CO-162 Blacksheep Staff` from being repopulated with all `CO-162@cowg.cap.gov` members during CAPWATCH unit-prefix sync.

When creating a missing Team, the script uses `Commanders.txt` to find the CAPWATCH commander for that unit and resolves the commander's CAPID to a Microsoft Graph user by `employeeId`. The commander is included as the initial Team owner because Microsoft Graph requires an owner when creating a Team in application context. After the Team is provisioned, the same run syncs members from the unit distribution group, such as `CO-200@cowg.cap.gov`, into the new Team. If the commander row or matching Graph user cannot be found, the script logs the issue and skips Team creation for that unit.

Manual guest Team assignments are read from the project root file `CAPWATCHpwsh/ManualGuests.csv`. Use the `Teams` column for one or more exact Team display names separated by semicolons, for example `"COWG Winter Encampment;CO-200 Example Squadron"`. Use the optional `Sponsor` column to record who approved the account; `updateTeams` includes it in assignment logs. Run `checkAccounts` first so the B2B guest exists; then `updateTeams` can resolve the user by manual `employeeId` or email and add them to the Team.

CAC Team assignments are read from CAPWATCH duty-position files. Cadets in `CadetDutyPositions.txt` with `FunctArea` of `CAC Appointment` or a duty name containing `CAC` are synced as members of the `CAC` Team. Senior members in `DutyPosition.txt` with the duty `Cadet Advisory Council Senior Advisor` are synced as Team members and owners. The CAC sync enforces membership from CAPWATCH, so users who no longer hold a CAC appointment or senior advisor position are removed from the Team. Owner removal is skipped if no senior advisor owner can be resolved, which avoids leaving the Team without a CAPWATCH owner when source data or Graph resolution is incomplete.

### Manual Guest Team Assignment Flow

1. Add or update the guest row in `CAPWATCHpwsh/ManualGuests.csv`.
2. Run `checkAccounts` first so the Microsoft Entra B2B guest exists.
3. Run `updateTeams` with `EXECUTE` unset or `EXECUTE=false` to preview the Team assignment.
4. Set `EXECUTE=true` and run `updateTeams` to add the guest to the exact Team display name(s) listed in the `Teams` column.

When a sponsor is present, dry-run and execute logs include it in the assignment detail, for example:

```text
Would add manual guest to Team - kaitlin.reinhart@utwg.cap.gov -> COWG Winter Encampment (Sponsor: Stephanie.Moffitt@cowg.cap.gov)
```

---

## Error Handling

- The script stops execution if CAPWATCH data is stale (older than 40 hours).
- Errors during API calls or Team updates are logged for troubleshooting.

---

## Notes

- Ensure CAPWATCH data is up-to-date before running the script.
- The script assumes CAPWATCH data is stored in the `$($env:HOME)\data\CAPWatch` directory.
- The script uses the beta endpoint of Microsoft Graph API, which may be subject to changes.

---

## License

This project is licensed under the [MIT License](LICENSE).
