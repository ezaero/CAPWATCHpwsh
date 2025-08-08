# CAPWATCHSyncPWSH

## Overview

**CAPWATCHSyncPWSH** is a PowerShell-based automation toolkit for synchronizing CAP membership data from CAPWATCH with Microsoft Teams and Exchange Online. It leverages Microsoft Graph API and Azure Managed Identity to automate the creation, update, and management of Teams, users, and mail contacts based on authoritative CAPWATCH data.

This toolkit is designed to work with any CAP Wing and can be easily configured for your specific wing's requirements.

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
  - `User.ReadWrite.All`
  - `Directory.ReadWrite.All`
  - `Mail.Send` (if using email notifications)
- **CAPWATCH Data**: Configure your Azure Function to read from Azure File Storage or place the latest CAPWATCH CSV files in the `$($env:HOME)/data/CAPWatch` directory.
- **Wing Configuration**: Set up environment variables for your specific wing (see [Configuration Guide](CONFIGURATION.md) for details).

---

## Installation & Deployment

This toolkit is designed to be deployed as an Azure Function App. Follow these steps:

### 📋 Prerequisites
- Azure subscription with appropriate permissions
- PowerShell 7+ installed locally
- Azure CLI or Azure PowerShell module
- Git for cloning the repository

### 🚀 Quick Start

1. **Clone the repository:**
   ```bash
   git clone https://github.com/ezaero/CAPWATCHpwsh.git
   cd CAPWATCHpwsh
   ```

2. **Deploy infrastructure:**
   ```bash
   # Use Terraform to deploy Azure resources
   cd terraform
   terraform init
   terraform plan
   terraform apply
   ```

3. **Set up PowerShell modules:**
   ```powershell
   # Upload required modules to Azure Storage
   Connect-AzAccount
   ./Download_Modules.ps1
   ./Upload-ModulesToStorage.ps1 -StorageAccountName "your-storage-account" -ResourceGroup "your-rg"
   ```

4. **Deploy function app:**
   ```bash
   # Deploy the PowerShell code
   func azure functionapp publish your-function-app-name --powershell
   ```

### 📖 Detailed Documentation

- **[DEPLOYMENT.md](DEPLOYMENT.md)** - Complete deployment guide with step-by-step instructions
- **[MODULE-SETUP.md](MODULE-SETUP.md)** - Quick reference for PowerShell module setup
- **[CONFIGURATION.md](CONFIGURATION.md)** - Wing-specific configuration guide

### ⚡ Key Deployment Notes

- **Module Management**: Due to Azure Functions size limits, PowerShell modules are uploaded to Azure Storage and loaded at runtime
- **Hybrid Loading**: Uses both Azure Functions managed dependencies and custom storage-based loading for reliability
- **Deployment Size**: Optimized to ~53KB (down from 180MB+) through selective exclusions

---

## Usage

- Run the main scripts in the appropriate subfolders (`updateTeams`, `checkAccounts`, etc.) as needed.
- Review and update configuration or environment variables as required for your deployment.
- All logs will be written to the `$($env:HOME)/logs` directory.

---

## Project Structure

- `/updateTeams` – Synchronizes Microsoft Teams membership and ownership with CAPWATCH data for each unit.
- `/checkAccounts` – Creates, updates, and restores user accounts and mail contacts in Azure AD and Exchange based on CAPWATCH data.
- `/Maintenance` – Performs monthly cleanup: deletes expired member accounts and old log files.
- `/shared` – Provides shared utility functions, including logging and Microsoft Graph helpers.
- `/DLAnnouncements` – Manages distribution lists for CAP announcements, ensuring correct membership based on CAPWATCH data.
- `/DLOpsQuals` – Automates distribution group membership for operational qualifications (e.g., pilots, aircrew, ES) using CAPWATCH and OpsQuals data.
- `/DLSeniorsCadets` – Maintains distribution lists for senior and cadet members, updating group membership as CAPWATCH data changes.
- `/DLSpecTrack` – Tracks and manages specialty distribution lists (e.g., specific qualifications or roles) for targeted communications.
- `/download-extract-capwatch` – Handles downloading and extraction of CAPWATCH data files for use by other automation scripts.
- `/emailLogFile` – Sends log files or notifications via email to administrators for audit and troubleshooting purposes.

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
