# DLOpsQuals Script

## Overview

The `DLOpsQuals` PowerShell script is designed to manage and update the membership of various distribution groups in Microsoft Entra ID (Azure AD) based on CAPWATCH data. It ensures that the groups contain the correct members by synchronizing data from CAPWATCH files and Azure AD.

## Features

1. **Group Membership Management**:
   - Updates distribution groups based on CAPWATCH achievements and tasks.
   - Adds users to groups if they meet specific criteria.
   - Removes users from groups if they no longer meet the criteria.

2. **Azure Integration**:
   - Uses the Microsoft Graph API to retrieve and update group membership.
   - Authenticates securely using Azure Managed Identity.

3. **Logging**:
   - Logs all actions, including users added or removed from groups and any errors encountered.

## Prerequisites

- **Microsoft Graph PowerShell SDK**: Ensure the SDK is installed and authenticated before running the script.
- **Azure Permissions**:
  - `Group.ReadWrite.All`
  - `User.Read.All`
- **Azure Managed Identity**: The script uses a managed identity for authentication in production environments.
- **CAPWATCH Data**: Ensure the CAPWATCH CSV files (`MbrAchievements.txt` and `MbrTasks.txt`) are available in the specified directory.

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/your-repo/CAPWATCHSyncPWSH.git
   cd CAPWATCHSyncPWSH/DLOpsQuals

   ## Key Functions

### Compare-Arrays
Compares two arrays and identifies:
- Users present in both arrays.
- Users only in the first array (to be added).
- Users only in the second array (to be removed).

### GetGroupMemberIds
Retrieves the IDs of all current members of the specified group.

### ModifyGroupMembers
Adds or removes users from the group based on the comparison results.

---

## Logic Flow

### Retrieve All Users
- Fetches all users from Azure AD using the `GetAllUsers` function.

### Filter Users for Group Membership
- Filters users based on CAPWATCH achievements and tasks.
- Excludes users without an email address.

### Compare Membership
- Compares the filtered users with the current group members using the `Compare-Arrays` function.

### Update Group Membership
- Adds users to the group if they are not already members.
- Removes users from the group if they no longer meet the criteria.

### Logging
- Logs all actions, including users added or removed from groups and any errors encountered.

---

## Outputs

- **Logs**: All actions and errors are logged for auditing purposes.

---

## Error Handling

- Errors during user addition or removal are caught and logged for troubleshooting.

---

## Notes

- The script assumes CAPID is stored in the `officeLocation` property of Azure AD users.
- Ensure the CAPWATCH data is up-to-date before running the script.

---

## License

This project is licensed under the [MIT License](LICENSE).