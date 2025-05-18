# DLSpecTrack Script

## Overview

The `DLSpecTrack` PowerShell script is designed to manage and update the membership of specialty track distribution groups in Microsoft Entra ID (Azure AD). It ensures that the groups contain the correct members by synchronizing data from CAPWATCH files and Azure AD.

## Features

1. **Group Membership Management**:
   - Updates distribution groups for specialty tracks based on CAPWATCH data.
   - Adds users to groups if they meet specific criteria.
   - Handles the creation of distribution groups if they do not already exist.

2. **Azure Integration**:
   - Uses the Microsoft Graph API to retrieve and update group membership.
   - Authenticates securely using Azure Managed Identity.

3. **Logging**:
   - Logs all actions, including users added to groups and any errors encountered.

## Prerequisites

- **Microsoft Graph PowerShell SDK**: Ensure the SDK is installed and authenticated before running the script.
- **Azure Permissions**:
  - `Group.ReadWrite.All`
  - `User.Read.All`
- **Azure Managed Identity**: The script uses a managed identity for authentication in production environments.
- **CAPWATCH Data**: Ensure the CAPWATCH CSV file (`SpecTrack.txt`) is available in the specified directory.

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/your-repo/CAPWATCHSyncPWSH.git
   cd CAPWATCHSyncPWSH/DLSpecTrack

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

### Retrieve All Users
- Fetches all users from Azure AD using the `GetAllUsers` function.

### Retrieve Specialty Tracks
- Reads the `SpecTrack.txt` file to get a list of all specialty tracks.

### Filter Users for Group Membership
- Filters users based on their `officeLocation` (CAPID) and ensures they have a valid email address.

### Compare Membership
- Compares the filtered users with the current group members using the `Compare-Arrays` function.

### Update Group Membership
- Adds users to the group if they are not already members.

### Logging
- Logs all actions, including users added to groups and any errors encountered.

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

---

## License

This project is licensed under the [MIT License](LICENSE).