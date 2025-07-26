# DLAnnouncements Script

## Overview

The `DLAnnouncements` PowerShell script is designed to manage the membership of the "CO Wing Announcements" distribution group in Microsoft Entra ID (Azure AD). It ensures that the group contains the correct members based on specific criteria, such as user type and job title, by synchronizing data from Azure AD.

## Features

1. **Group Membership Management**:
   - Adds users to the "CO Wing Announcements" group if they meet the specified criteria.
   - Compares current group members with the desired membership list to identify users to add.

2. **Azure Integration**:
   - Uses the Microsoft Graph API to retrieve and update group membership.
   - Authenticates securely using Azure Managed Identity.

3. **Logging**:
   - Logs all actions, including users added to the group and any errors encountered.

## Prerequisites

- **Microsoft Graph PowerShell SDK**: Ensure the SDK is installed and authenticated before running the script.
- **Azure Permissions**:
  - `Group.ReadWrite.All`
  - `User.Read.All`
- **Azure Managed Identity**: The script uses a managed identity for authentication in production environments.

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/your-repo/CAPWATCHSyncPWSH.git
   cd CAPWATCHSyncPWSH/DLAnnouncements

## Key Functions

### Compare-Arrays
Compares two arrays and identifies:
- Users present in both arrays.
- Users only in the first array (to be added).

### GetGroupMemberIds
Retrieves the IDs of all current members of the specified group.

### ModifyGroupMembers
Adds users to the group based on the comparison results.

---

## Logic Flow

### Retrieve All Users
- Fetches all users from Azure AD using the `GetAllUsers` function.

### Filter Users for Group Membership
- Filters users based on the following criteria:
  - `employeeType` is `CADET` or `SENIOR`.
  - `jobTitle` contains `PARENT`.
  - Excludes users without an email address.

### Compare Membership
- Compares the filtered users with the current group members using the `Compare-Arrays` function.

### Update Group Membership
- Adds users to the group if they are not already members.

### Logging
- Logs all actions, including users added to the group and any errors encountered.

---

## Outputs

- **Logs**: All actions and errors are logged for auditing purposes.

---

## Error Handling

- Errors during user addition are caught and logged for troubleshooting.

---

## Notes

- The script does not remove users from the group if they are no longer in the filtered list. This is intentional to avoid removing users whose accounts may have been deleted.

---

## License

This project is licensed under the [MIT License](LICENSE).

