# CAPWATCHSyncPWSH

## Overview

**CAPWATCHSyncPWSH** is a PowerShell-based automation toolkit for synchronizing CAP membership data from CAPWATCH with Microsoft Teams and Exchange Online. It leverages Microsoft Graph API and Azure Managed Identity to automate the creation, update, and management of Teams, users, and mail contacts based on authoritative CAPWATCH data.

---

## How It Works in Azure Functions

- **Serverless Execution**: Runs as an Azure Function App, so no dedicated server or VM is required.
- **No Database Required**: The solution does not use or require a database. All data processing is performed in-memory at runtime.
- **Daily Data Refresh**: CAPWATCH data is downloaded daily into Azure File Storage using a scheduled process.
- **Real-Time Processing**: When triggered, the Azure Function loads the latest CAPWATCH CSV files directly from Azure File Storage, processes them in real time, and runs queries against Microsoft 365 (Entra ID/Azure AD and Exchange Online) using Microsoft Graph API.
- **Automation**: All synchronization, creation, and update operations are performed automatically based on the latest data, with no manual intervention required.

---

## Features

- **Microsoft Teams Automation**
  - Creates and updates Teams for each unit.
  - Synchronizes Team members and owners with CAPWATCH data.
  - Ensures correct aliases and ownership for each Team.

- **Exchange Online Integration**
  - Manages mail contacts for members and guests.
  - Removes or restores contacts based on membership status.

- **CAPWATCH Data Processing**
  - Reads and processes CAPWATCH CSV exports from Azure File Storage.
  - Filters and normalizes member data for downstream automation.

- **Logging & Error Handling**
  - Logs all actions and errors to a dedicated logs directory.
  - Stops execution if CAPWATCH data is stale.

- **Azure Integration**
  - Uses Azure Managed Identity for secure authentication.
  - Follows Azure and Microsoft Graph best practices for permissions and security.

---

## Prerequisites

- **Microsoft Graph PowerShell SDK** installed and available in your environment.
- **Azure Function App** (or automation host) with Managed Identity enabled and granted the following Microsoft Graph API permissions (with admin consent):
  - `Group.ReadWrite.All`
  - `TeamMember.ReadWrite.All`
  - `User.Read.All`
  - `Directory.Read.All`
- **CAPWATCH Data**: Place the latest CAPWATCH CSV files (`Organization.txt`, `Commanders.txt`, `DutyPosition.txt`, etc.) in the `$($env:HOME)/data/CAPWatch` directory or configure your Azure Function to read from Azure File Storage.

---

## Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/your-repo/CAPWATCHSyncPWSH.git
   cd CAPWATCHSyncPWSH
   ```

2. **Configure your environment:**
   - Ensure all required PowerShell modules are installed.
   - Set up your Azure Function App or automation environment with the correct permissions.

3. **Prepare CAPWATCH data:**
   - Place the required CSV files in the data directory as described above, or ensure your Azure File Storage is populated daily.

---

## Usage

- Run the main scripts in the appropriate subfolders (`updateTeams`, `checkAccounts`, etc.) as needed.
- Review and update configuration or environment variables as required for your deployment.
- All logs will be written to the `$($env:HOME)/logs` directory.

---

## Project Structure

- `/updateTeams` – Scripts for managing and synchronizing Microsoft Teams.
- `/checkAccounts` – Scripts for managing Exchange contacts and user accounts.
- `/shared` – Shared utility functions and logging.
- `/output` – Output and export files (excluded from version control).
- `/data` – CAPWATCH data files (excluded from version control).

---

## Security & Best Practices

- **Do not commit secrets or credentials.** Use environment variables or Azure Key Vault for sensitive data.
- **Review all scripts for organization-specific information** before making the repository public.
- **Follow [Azure best practices](https://learn.microsoft.com/azure/cloud-adoption-framework/ready/azure-best-practices/)** for automation and security.
- **Use a `.gitignore`** to exclude logs, output, credentials, and IDE files.

---

## License

This project is licensed under the [MIT License](LICENSE).

---

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for improvements or bug fixes.

---

## Disclaimer

This project is provided as-is and is not officially supported by Civil Air Patrol or Microsoft. Use at your own risk.
