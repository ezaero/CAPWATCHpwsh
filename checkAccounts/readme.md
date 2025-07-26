# CAPWATCHSyncPWSH - Synchronize CAPWATCH Data with Microsoft Entra ID (Azure AD)

## Overview

This PowerShell script synchronizes CAPWATCH data with Microsoft Entra ID (Azure AD) and ensures accurate user information in Office 365. It automates the process of managing user accounts, including creating, updating, and restoring users based on CAPWATCH data. Account deletion for expired members is handled separately by the Maintenance function.

## Features

1. **Data Synchronization**:
   - Imports CAPWATCH data from CSV files (`MbrContact.txt`, `Member.txt`, and `DutyPosition.txt`).
   - Combines and processes the data to create a unified dataset for user management.

2. **User Account Creation and Restoration**:
   - Identifies users to be added as Office 365 guest accounts if missing.
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
- **CAPWATCH Data**: Ensure the CAPWATCH CSV files (`MbrContact.txt`, `Member.txt`, and `DutyPosition.txt`) are available in the specified directory.

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

- The script assumes CAPID is stored in the `officeLocation` property of Azure AD users.
- Ensure the CAPWATCH data is up-to-date before running the script.

---

## License

This project is licensed under the [MIT License](LICENSE).