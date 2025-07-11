# Specialty Track Distribution Group Automation

## Overview

This script automates the management of Microsoft 365 distribution groups for Civil Air Patrol specialty tracks. It reads CAPWATCH specialty track data, queries Microsoft Entra ID (Azure AD), and ensures each specialty track has a corresponding distribution group with the correct members.

---

## Features

- **Automated Group Management**: Creates distribution groups for each specialty track if they do not exist.
- **Membership Synchronization**: Adds users to groups based on CAPWATCH data and removes them if their account is deleted.
- **Logging**: Logs all actions and errors for auditing and troubleshooting.
- **No Database Required**: All data is processed in-memory from CAPWATCH CSV files and Microsoft 365 queries.

---

## How It Works

- Runs as an Azure Function or PowerShell automation (no database required).
- Loads the latest `SpecTrack.txt` from Azure File Storage or local data directory (CAPWATCH data is refreshed daily).
- Queries Microsoft Entra ID (Azure AD) for user information in real time.
- For each specialty track, ensures a distribution group exists and synchronizes its membership.

---

## Prerequisites

- Microsoft Graph PowerShell SDK
- ExchangeOnlineManagement PowerShell module
- CAPWATCH `SpecTrack.txt` file in the data directory or Azure File Storage
- Azure Function App or automation host with Managed Identity and required Microsoft Graph/Exchange permissions

---

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/your-repo/CAPWATCHSyncPWSH.git
   cd CAPWATCHSyncPWSH/DLSpecTrack
   ```
2. Place the latest `SpecTrack.txt` in the data directory or configure Azure File Storage for daily refresh.
3. Ensure all required modules and permissions are in place.

---

## Usage

- Run the script in Azure Functions or locally with the required modules and permissions.
- Review logs in the `logs/` directory for results and troubleshooting.

---

## Key Functions

### Compare-Arrays
- Compares two arrays and identifies:
  - Users present in both arrays.
  - Users only in the first array (to be added).
  - Users only in the second array (to be removed).

### GetGroupMemberIds
- Retrieves the IDs of all current members of the specified group.
- Creates the group if it does not already exist.

### ModifyGroupMembers
- Adds users to the group if they are not already members.
- Does not remove users from the group to avoid unintended deletions.

---

## Logic Flow

1. **Retrieve All Users**: Fetches all users from Azure AD using the `GetAllUsers` function.
2. **Retrieve Specialty Tracks**: Reads the `SpecTrack.txt` file to get a list of all specialty tracks.
3. **Filter Users for Group Membership**: Filters users based on their `officeLocation` (CAPID) and ensures they have a valid email address.
4. **Compare Membership**: Compares the filtered users with the current group members using the `Compare-Arrays` function.
5. **Update Group Membership**: Adds users to the group if they are not already members.
6. **Logging**: Logs all actions, including users added to groups and any errors encountered.

---

## Outputs

- **Logs**: All actions and errors are logged for auditing purposes.

---

## Error Handling

- Errors during user addition are caught and logged for troubleshooting.

---

## Notes

- The script assumes CAPID is stored in the `officeLocation` property of Azure AD users.
- Ensure the CAPWATCH data is up-to-date before running the script.
- No database is required; all processing is done in-memory and in real time.

---

## Security & Best Practices

- Do not commit secrets or credentials.
- Use environment variables or Azure Key Vault for sensitive data.
- Review scripts for organization-specific information before making public.

---

## License

This project is licensed under the [MIT License](../LICENSE).