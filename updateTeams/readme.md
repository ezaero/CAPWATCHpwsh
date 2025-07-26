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

3. **Owner Management**:
   - Ensures the unit commander is the owner of the Team.
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