# Email Log File Script

## Overview

The `emailLogFile` PowerShell script is designed to send a daily log file via email using the Microsoft Graph API. This script is executed as part of an Azure Function App and ensures that the log file for the current day is sent to a specified recipient.

## Features

1. **Email Log File**:
   - Sends the daily log file as an email attachment.
   - Uses Microsoft Graph API for secure and reliable email delivery.

2. **Azure Integration**:
   - Authenticates with Microsoft Graph API using Azure Managed Identity.
   - Retrieves an access token using the `Get-AzAccessToken` cmdlet.

3. **Error Handling**:
   - Checks if the log file exists before attempting to send the email.
   - Logs all actions and errors for auditing purposes.

---

## Prerequisites

- **Microsoft Graph PowerShell SDK**:
  - Ensure the SDK is installed and authenticated before running the script.

- **Azure Managed Identity**:
  - The Azure Function App must have permissions to send emails via Microsoft Graph API.

- **Log File Directory**:
  - Ensure the log files are stored in the `$($env:HOME)\logs` directory.

---

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/your-repo/CAPWATCHSyncPWSH.git
   cd CAPWATCHSyncPWSH/emailLogFile

## Key Components

### Microsoft Graph API Integration
- Authenticates with Microsoft Graph API using an access token retrieved via `Get-AzAccessToken`.
- Sends the email using the `Invoke-MgGraphRequest` cmdlet.

### Log File Handling
- Checks if the log file for the current day exists in the `$($env:HOME)\logs` directory.
- Reads the log file content and encodes it in Base64 format for email attachment.

### Email Payload
- Constructs the email payload with the following details:
  - **To Address**: The recipient's email address.
  - **From Address**: The sender's email address.
  - **Subject**: "Daily Log File".
  - **Body**: A brief message with the current date.
  - **Attachment**: The daily log file.

---

## Logic Flow

### Retrieve Access Token
- Fetches an access token for Microsoft Graph API using Azure Managed Identity.

### Check Log File
- Verifies if the log file for the current day exists in the `$($env:HOME)\logs` directory.

### Read and Encode Log File
- Reads the log file content and encodes it in Base64 format.

### Construct Email Payload
- Creates the email payload with the subject, body, recipient, and attachment.

### Send Email
- Sends the email using the Microsoft Graph API.

### Logging
- Logs all actions, including email sending and any errors encountered.

---

## Outputs

- **Logs**: All actions and errors are logged for auditing purposes.

---

## Error Handling

- The script stops execution if the log file does not exist.
- Errors during email sending are logged for troubleshooting.

---

## Notes

- Ensure the sender's email address is valid and authorized to send emails via Microsoft Graph API.
- The script assumes the log file is named in the format `script_log_yyyy-MM-dd.txt`.

---

## License

This project is licensed under the [MIT License](LICENSE).