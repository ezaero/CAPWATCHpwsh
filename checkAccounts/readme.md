# CAPWATCHSyncPWSH - Synchronize CAPWATCH Data with Microsoft Entra ID (Azure AD)

## Overview

This PowerShell script synchronizes CAPWATCH data with Microsoft Entra ID (Azure AD) and ensures accurate user information in Office 365. It automates the process of managing user accounts, including creating, updating, and restoring users based on CAPWATCH data. Account deletion for expired members is handled separately by the Maintenance function.

## Features

1. **Data Synchronization**:
   - Imports CAPWATCH data from CSV files (`MbrContact.txt`, `Member.txt`, `DutyPosition.txt`, and `Commanders.txt`).
   - Combines and processes the data to create a unified dataset for user management.

2. **User Account Creation and Restoration**:
   - Identifies users to be added as Office 365 guest accounts if missing.
   - Imports optional one-off external users from `ManualGuests.csv` and onboards them through the same guest invitation path.
   - Restores deleted accounts if a user renews their membership.
   - Adds new mail contacts for Aerospace Education Members (AEMs).

3. **User Account Updates**:
   - Updates existing Office 365 accounts with CAPID, grade (promotion), duty positions, unit assignment, and email address.
   - Ensures all Office 365 accounts have the correct CAPID, duty positions, and unit information.
   - Maintains accurate department, job title, and display name information for all users.

4. **Error Handling and Logging**:
   - Logs all actions and errors for auditing purposes.
   - Exports users with missing CAPIDs and identifies duplicate display names.

5. **Azure Integration**:
   - Connects to Microsoft Graph API using the Microsoft Graph PowerShell SDK.
   - Uses Azure Managed Identity for secure authentication.

6. **Account Deletion**:
   - Account deletion for expired members and cleanup of stale accounts is performed by the Maintenance function, which runs monthly.

## Prerequisites

- **Microsoft Graph PowerShell SDK**: Ensure the SDK is installed and authenticated before running the script.
- **Azure Permissions**:
  - `User.Read.All`
  - `User.ReadWrite.All`
  - `Directory.ReadWrite.All`
- **CAPWATCH Data**: Ensure the CAPWATCH CSV files (`MbrContact.txt`, `Member.txt`, `DutyPosition.txt`, and `Commanders.txt`) are available in the specified directory.
- **Manual Guests**: To onboard non-CAPWATCH users, place `ManualGuests.csv` in the project root (`CAPWATCHpwsh/ManualGuests.csv`) with columns such as `Email`, `DisplayName`, `NameFirst`, `NameLast`, `Grade`, `Type`, `Unit`, `Teams`, `Sponsor`, `EmployeeId`, and `Enabled`.

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/your-repo/CAPWATCHSyncPWSH.git
   cd CAPWATCHSyncPWSH

## Key Functions

### Compare-Arrays
Compares two arrays and identifies:
- Users present in both arrays.
- Users only in the first array (to be added).
- Users only in the second array (to be removed).

### Combine
Combines data from the `Member.txt` and `MbrContact.txt` files into a unified dataset.

### DutyPositions
Processes the `DutyPosition.txt` file and creates a hashtable of duty positions for each CAPID.

### GetAllUsers
Retrieves all users from Microsoft Graph API.

### AddNewGuest
Creates a new guest user in Azure AD.

### RestoreDeletedAccounts
Restores deleted accounts if a user renews their membership.

---

## Manual Guest Onboarding

Use `ManualGuests.csv` for one-off users who need COWG Microsoft 365/Teams access but are not present in CAPWATCH, such as out-of-wing encampment staff.

### File Location

Place the file at the project root:

```text
CAPWATCHpwsh/ManualGuests.csv
```

### CSV Columns

```csv
Email,DisplayName,NameFirst,NameLast,Grade,Type,Unit,Teams,Sponsor,EmployeeId,Enabled
```

- `Email`: External user's email address. Required.
- `DisplayName`: Full display name. Used to derive first/last name if those columns are blank.
- `NameFirst`, `NameLast`: Preferred first and last name values.
- `Grade`: Display role/title, such as `Winter Encampment Staff`.
- `Type`: Use `MANUAL` for non-CAPWATCH one-off accounts.
- `Unit`: Local grouping label, such as `WINTER-ENCAMPMENT`.
- `Teams`: One or more exact Team display names separated by semicolons. Team assignment is performed by `updateTeams`, not `checkAccounts`.
- `Sponsor`: COWG sponsor/approver for the account. This is retained in CSV-derived records, logs, and notification output, but is not written to Microsoft Entra ID.
- `EmployeeId`: Optional. If blank, the script creates a stable `MANUAL-<hash>` ID from the email address.
- `Enabled`: Set to `TRUE` or leave blank to process the row. Set to `FALSE`, `NO`, `0`, or `DISABLED` to ignore the row.

Example:

```csv
Email,DisplayName,NameFirst,NameLast,Grade,Type,Unit,Teams,Sponsor,EmployeeId,Enabled
kaitlin.reinhart@utwg.cap.gov,Kaitlin Reinhart,Kaitlin,Reinhart,Winter Encampment Staff,MANUAL,WINTER-ENCAMPMENT,"COWG Winter Encampment",Stephanie.Moffitt@cowg.cap.gov,,TRUE
```

### Dry Run

Leave `EXECUTE` unset or set `EXECUTE=false`, then run `checkAccounts`. In dry-run mode the function logs what it would do, but it does not create the guest account or send the B2B invitation.

Expected dry-run log signals include:

```text
Loaded 1 manual guest record(s) from .../ManualGuests.csv
[DRY-RUN] Guest user creation skipped (would create and send invitation)
```

### Execute

Set `EXECUTE=true` and run `checkAccounts` to create the B2B guest, send the invitation, and stamp the manual `employeeId`, unit, title, type, and mail metadata.

After the guest exists, run `updateTeams` to add the user to the Team(s) listed in the `Teams` column. `updateTeams` also supports dry-run mode, so you can preview Team assignment before applying it.

### Removing Access

To stop processing a manual guest, set `Enabled=FALSE` or remove the row. While the row is enabled, monthly maintenance treats the manual `employeeId` as valid and will not delete the account just because it is absent from CAPWATCH.

## Outputs

- **Logs**: All actions and errors are logged for auditing purposes.
- **Exports**:
  - Users with missing CAPIDs are exported to `noCAPID.csv`.
  - Duplicate display names are logged for review.

---

## Error Handling

- The script aborts execution if CAPWATCH data is stale (older than 48 hours).
- Errors during user creation, updates, or deletions are logged for troubleshooting.

---

## Notes

- The script resolves CAPID from the beta Microsoft Graph `employeeId` property. `Commanders.txt` can override a member's unit when correcting `companyName` for assigned commanders.
- Manual guest rows receive a stable `employeeId` in the form `MANUAL-<hash>` unless `EmployeeId` is supplied. Keep the row enabled while the guest should remain active; monthly maintenance treats enabled manual guest IDs as valid accounts.
- The optional `Sponsor` column records who approved the one-off account. It is retained in the CSV-derived record and notification/log output, but is not written into Microsoft Entra ID.
- The `DoNotContact` flag does not suppress account attribute maintenance such as CAPID, unit, rank, or duty-position updates.
- Ensure the CAPWATCH data is up-to-date before running the script.

---

## License

This project is licensed under the [MIT License](LICENSE).
