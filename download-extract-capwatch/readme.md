
# Download and Extract CAPWATCH Script

# SYNOPSIS

Downloads data from CAPWATCH and extracts it to a folder accesible to the Azure Function App

# DESCRIPTION

This script executes on a timer as part of an Azure Function App and is responsible for
downloading and extracting data retrieved from CAPWATCH for Colorado Wing CAP. Data
downloaded and extracted by this script will reside in the $($env:HOME)\data\CAPWatch
directory for use by other scripts in this Azure Function App

# NOTES

This script pulls credentials for CAPWATCH from an Azure Key Vault specified in the $KeyVaultName variable.
At the time of writing this script, the Key Vault (cowgcapwatch) is set to allow members of the IT staff to
write secrets but not read them. This Function App is the only resource with permissions to retrieve secret
values in an effort to protect the personal credentials of the user-account tied to the CAPWATCH download
API as they are the same credentials used to log-in to eServices.

## Overview

The `download-extract-capwatch` PowerShell script is designed to download and extract CAPWATCH data for the Colorado Wing CAP. This script is executed as part of an Azure Function App on a timer trigger. The downloaded data is extracted to a directory accessible by other scripts within the Azure Function App.

## Features

1. **CAPWATCH Data Retrieval**:
   - Downloads CAPWATCH data from the CAPNHQ API using credentials securely stored in Azure Key Vault.
   - Extracts the downloaded data into a specified directory for further processing.

2. **Azure Integration**:
   - Retrieves credentials (`capwatch-username` and `capwatch-password`) from Azure Key Vault.
   - Uses Azure Managed Identity to securely access the Key Vault.

3. **Error Handling**:
   - Stops execution on errors to ensure data integrity.
   - Logs all actions and errors for auditing purposes.

4. **File Management**:
   - Deletes any existing CAPWATCH ZIP file before downloading a new one.
   - Extracts the downloaded ZIP file to the `$($env:HOME)\data\CAPWatch` directory.

---

## Prerequisites

- **Azure Key Vault**:
  - Ensure the Key Vault contains the secrets `capwatch-username` and `capwatch-password`.
  - The Azure Function App must have permissions to retrieve secrets from the Key Vault.

- **Azure PowerShell Module**:
  - Ensure the `Az` PowerShell module is installed and authenticated.

- **CAPWATCH API Access**:
  - The CAPWATCH API credentials must be valid and authorized to access the required data.

---

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/your-repo/CAPWATCHSyncPWSH.git
   cd CAPWATCHSyncPWSH/download-extract-capwatch

## Key Components

### Azure Key Vault Integration
- The script retrieves CAPWATCH API credentials (`capwatch-username` and `capwatch-password`) from Azure Key Vault using the `Get-AzKeyVaultSecret` cmdlet.

### CAPWATCH Data Download
- Downloads the CAPWATCH ZIP file from the CAPNHQ API using the retrieved credentials.

### Data Extraction
- Extracts the downloaded ZIP file to the `$($env:HOME)\data\CAPWatch` directory.

---

## Logic Flow

### Retrieve Credentials
- Fetches the CAPWATCH API credentials from Azure Key Vault.

### Delete Existing Files
- Deletes any existing CAPWATCH ZIP file to ensure a clean download.

### Download CAPWATCH Data
- Downloads the CAPWATCH ZIP file from the CAPNHQ API using the retrieved credentials.

### Extract Data
- Extracts the downloaded ZIP file to the `$($env:HOME)\data\CAPWatch` directory.

### Logging
- Logs all actions, including file deletions, downloads, and extractions.

---

## Outputs

- **Logs**: All actions and errors are logged for auditing purposes.

---

## Error Handling

- The script stops execution on any error to ensure data integrity.
- Errors during credential retrieval, file download, or extraction are logged for troubleshooting.

---

## Notes

- The script assumes the CAPWATCH API credentials are securely stored in Azure Key Vault.
- Ensure the Azure Function App has the necessary permissions to access the Key Vault and execute the script.

---

## License

This project is licensed under the [MIT License](LICENSE).