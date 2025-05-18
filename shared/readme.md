# Shared Functions Script

## Overview

The `shared.ps1` PowerShell script contains reusable functions that support various operations within the CAPWATCHSyncPWSH project. These functions handle logging, user retrieval, and other shared tasks, ensuring consistency and reducing redundancy across scripts.

---

## Features

1. **Logging**:
   - Logs messages to a file and outputs them to the console.
   - Ensures log directories are created if they do not exist.

2. **Microsoft Graph API Integration**:
   - Retrieves all users from Microsoft Graph API.
   - Retrieves deleted users from Microsoft Graph API.

3. **Error Handling**:
   - Logs errors encountered during API calls for troubleshooting.

---

## Functions

### `Write-Log`
- **Purpose**: Logs messages to a file and outputs them to the console.
- **Parameters**:
  - `Message` (string): The message to log.
- **Behavior**:
  - Ensures the log directory exists.
  - Writes the log message with a timestamp to the log file and console.
- **Log File Location**: `$($env:HOME)\logs\script_log_yyyy-MM-dd.txt`

### `GetAllUsers`
- **Purpose**: Retrieves all users from Microsoft Graph API.
- **Parameters**:
  - `SelectFields` (string): A comma-separated list of fields to retrieve for each user. Default: `"mail,displayName,officeLocation,companyName,employeeId,id,employeeType,jobTitle"`.
- **Behavior**:
  - Fetches users in a paginated manner using the `@odata.nextLink` property.
  - Logs errors if the API call fails.
- **Returns**: An array of user objects.

### `GetDeletedUsers`
- **Purpose**: Retrieves deleted users from Microsoft Graph API.
- **Behavior**:
  - Queries the `directory/deletedItems/microsoft.graph.user` endpoint.
  - Fetches deleted users in a paginated manner using the `@odata.nextLink` property.
- **Returns**: An array of deleted user objects.

---

## Prerequisites

- **Microsoft Graph PowerShell SDK**:
  - Ensure the SDK is installed and authenticated before using these functions.
- **Azure Permissions**:
  - The script requires permissions to access Microsoft Graph API endpoints:
    - `User.Read.All`
    - `Directory.Read.All`

---

## Usage

### Importing the Script
To use the shared functions in another script, include the following line:
```powershell
. "$PSScriptRoot\..\shared\shared.ps1"