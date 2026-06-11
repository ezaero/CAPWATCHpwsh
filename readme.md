
# CAPWATCHSyncPWSH

[![Watch the CAPWATCHSync Overview Video](https://img.shields.io/badge/Watch%20Video-Overview-blue?style=for-the-badge&logo=azurefunctions)](https://capwatchpwshtraining.blob.core.windows.net/training/CAPWATCHSync.mp4)

**[Watch a short video overview of CAPWATCHSyncPWSH features, benefits, and architecture.](https://capwatchpwshtraining.blob.core.windows.net/training/CAPWATCHSync.mp4)**

## Overview

**CAPWATCHSyncPWSH** is a PowerShell-based automation toolkit for synchronizing CAP membership data from CAPWATCH with Microsoft Teams and Exchange Online. It leverages Microsoft Graph API and Azure Managed Identity to automate the creation, update, and management of Teams, users, and mail contacts based on authoritative CAPWATCH data.

**This toolkit is designed to work with any CAP Wing** and has been successfully deployed across multiple Wings. It can be easily configured for your specific wing's requirements with just a few configuration changes.

## Table of Contents

- [Quick Start for New Wings](#quick-start-for-new-wings)
- [How It Works](#how-it-works)
- [Features](#features)
- [Runtime Prerequisites](#runtime-prerequisites)
- [Installation & Deployment](#installation--deployment)
- [Post-Deployment Configuration](#post-deployment-configuration)
- [Usage](#usage)
- [Project Structure](#project-structure)
- [Troubleshooting](#troubleshooting)
- [Azure Cost Estimation](#azure-cost-estimation)
- [Frequently Asked Questions](#frequently-asked-questions)
- [Multi-Wing Collaboration](#multi-wing-collaboration)
- [Contributing](#contributing)
- [Support and Contact](#support-and-contact)

---

## How It Works

### Architecture Overview

**CAPWATCHSyncPWSH** runs as a serverless Azure Function App with the following architecture:

- **Serverless Execution**: Runs on Azure Functions Consumption Plan - no dedicated servers required
- **Scheduled Automation**: Functions execute on configurable timers (daily, hourly, weekly)
- **No Database Required**: All data processing is performed in-memory using CAPWATCH CSV exports
- **Secure Authentication**: Uses Azure Managed Identity for all API access (Microsoft Graph, Exchange Online)

### Data Flow

1. **CAPWATCH Data Download** (`download-extract-capwatch`)
   - Scheduled function downloads latest CAPWATCH data daily
   - Authenticates using credentials from Azure Key Vault
   - Extracts and saves CSV files to Azure Storage (`$HOME/data/CAPWatch`)
   - Validates data freshness before proceeding

2. **Data Processing & Synchronization**
   - Functions load CSV data from Azure Storage
   - Process membership, unit, and qualification data
   - Query current state from Microsoft 365 using Microsoft Graph API
   - Calculate differences and required changes

3. **Automated Updates**
   - **Teams Management**: Creates/updates Teams for units, syncs membership and ownership
   - **User Accounts**: Creates, updates, or restores user accounts in Azure AD
   - **Distribution Lists**: Manages Exchange Online distribution groups for announcements, qualifications, and specialty tracks
   - **Mail Contacts**: Maintains mail contacts for external members

4. **Logging & Monitoring**
   - All operations logged to `$HOME/logs` directory
   - Application Insights captures telemetry and errors
   - Optional email notifications for administrators

### Multi-Wing Deployment

**This solution is designed to support multiple CAP Wings independently:**

- Each Wing deploys their own isolated Azure infrastructure
- Configuration is parameterized by Wing designator (2-letter code)
- CAPWATCH data is wing-specific using Organization ID
- No cross-wing dependencies or shared resources
- Each Wing maintains their own credentials and permissions

**Currently deployed Wings:** Colorado (CO), and expanding to additional Wings

---

## Quick Start for New Wings

Ready to deploy this solution for your Wing? Follow this checklist:

- [ ] **Review the video overview** to understand the solution
- [ ] **Gather required information**:
  - Wing 2-letter designator
  - CAPWATCH Organization ID
  - Exchange Online organization domain
  - CAPWATCH credentials (eServices username/password)
- [ ] **Prepare your environment**:
  - Azure subscription access (Contributor role)
  - Azure AD admin access (Application Administrator role)
  - Install PowerShell 7.6+, Azure CLI, and Terraform
- [ ] **Clone the repository** and navigate to terraform directory
- [ ] **Configure terraform.tfvars** with your Wing's values
- [ ] **Deploy infrastructure** with Terraform
- [ ] **Grant API permissions** in Azure AD (admin consent)
- [ ] **Upload PowerShell modules** to Azure Storage
- [ ] **Deploy function code** with Azure Functions Core Tools
- [ ] **Test and validate** starting with data download
- [ ] **Configure schedules** and enable production mode

**Estimated deployment time:** 2-3 hours for first-time deployment

---

## Multi-Wing Deployment Structure

This repository uses a **configuration-driven approach** to support multiple CAP Wings from a single codebase:

### Repository Organization

```
CAPWATCHpwsh/
├── shared/                    # Shared modules used by all wings
├── wings/                     # Wing-specific configurations
│   ├── colorado/              # Colorado Wing
│   │   ├── local.settings.json
│   │   └── terraform.tfvars
│   └── montana/               # Montana Wing
│       ├── local.settings.json
│       └── terraform.tfvars
├── [function folders]/        # Shared Azure Functions (all wings)
└── terraform/                 # Infrastructure templates (parameterized)
```

### Deployment Strategy

Each wing has:
- **Separate Azure subscription** or resource group
- **Wing-specific configuration files** (stored in `/wings/{wing}/`)
- **Independent Function App** and supporting Azure resources
- **Wing-specific Cosmos DB databases** for data isolation
- **Shared codebase** for all logic and functions

### Starting Your Wings Deployment

1. **For local development**: Copy the wing's config to root directory
   ```bash
   cp wings/{wing}/local.settings.json ./local.settings.json
   ```

2. **For infrastructure as code**: Use the wing's terraform variables
   ```bash
   terraform apply -var-file="../wings/{wing}/terraform.tfvars"
   ```

3. **To add a new wing**: See the [Multi-Wing Deployment Guide](wings/WING-DEPLOYMENT.md)

**Benefits of this approach:**
- ✅ Bug fixes and improvements benefit all wings automatically
- ✅ Shared code stays in sync
- ✅ Complete independence for each wing's resources
- ✅ Easy to add new wings without code duplication

For detailed multi-wing setup instructions, see [WING-DEPLOYMENT.md](wings/WING-DEPLOYMENT.md).

---

## Features

### Microsoft Teams Automation
- **Automatic Team Creation**: Creates Teams for each CAP unit based on CAPWATCH data
- **Membership Synchronization**: Adds/removes members based on current CAPWATCH membership
- **Owner Management**: Assigns unit commanders and deputy commanders as Team owners
- **Team Naming**: Standardized naming using Wing designator and unit number (e.g., `CO-022 Boulder Composite`)
- **Alias Management**: Sets proper Team email aliases aligned with unit structure

### User Account Management
- **Account Creation**: Automatically creates Azure AD accounts for new members
- **Profile Updates**: Updates display names, ranks, and attributes from CAPWATCH data
- **Account Restoration**: Restores soft-deleted accounts when members rejoin
- **Mail Contacts**: Creates and manages mail contacts for external/guest members
- **License Management**: Tracks and manages Microsoft 365 license assignments

### Distribution List Management
- **Announcements Lists**: Wing-level, region-level, and unit-level announcement distribution
- **Operational Qualifications**: Manages lists for pilots, aircrew, mission staff, ground team, and ES personnel
- **Senior/Cadet Lists**: Separate distribution groups for senior members and cadets
- **Specialty Tracks**: Custom distribution lists for specific programs or qualifications
- **Dynamic Updates**: All lists automatically updated as CAPWATCH data changes

### CAPWATCH Data Integration
- **Automated Downloads**: Scheduled daily downloads of CAPWATCH CSV exports
- **Data Validation**: Checks for data freshness and completeness
- **In-Memory Processing**: No database required - processes CSV files directly
- **Multi-File Support**: Handles member, unit, operational qualifications, and other CAPWATCH exports
- **Error Detection**: Validates data integrity before processing changes

### Logging & Monitoring
- **Comprehensive Logging**: All operations logged to dedicated log files
- **Application Insights**: Centralized telemetry and error tracking
- **Email Notifications**: Optional daily log summaries sent to administrators
- **Audit Trail**: Complete history of all changes made to accounts and groups
- **Error Handling**: Graceful failure handling with detailed error messages

### Security & Compliance
- **Azure Managed Identity**: No hardcoded credentials - uses managed identity for all API access
- **Key Vault Integration**: CAPWATCH credentials stored securely in Azure Key Vault
- **Least Privilege**: Minimal required permissions following Azure best practices
- **Secure Module Loading**: PowerShell modules loaded from secure Azure Storage
- **No Plaintext Secrets**: All sensitive data encrypted at rest and in transit

### Safety & Testing Features
- **🔍 Dry-Run Mode** (Default): Preview all changes without modifying anything
- **Progressive Rollout**: Test on one wing, safely deploy to others
- **Change Logging**: Clear logs showing what would/will change before execution
- **Reversible**: All operations logged and can be reviewed for audit

### Azure Integration
- **Serverless Architecture**: Runs on Azure Functions Consumption Plan (pay-per-use)
- **Automatic Scaling**: Scales automatically based on workload
- **Timer-Based Execution**: Scheduled functions run automatically without manual intervention
- **Infrastructure as Code**: Complete Terraform configuration for reproducible deployments
- **Multi-Wing Support**: Parameterized configuration for any CAP Wing

---

## Dry-Run Mode (Safety-First Deployment)

🔍 **All functions run in safe dry-run mode by default!**

This means functions will:
- ✅ Query and analyze data
- ✅ Log what changes would be made
- ❌ NOT create Teams, modify users, or send emails

**To preview changes without risk:** Just run the functions with the default settings.

**To apply changes:** Set the `EXECUTE=true` environment variable in your Function App configuration.

**→ See [DRY-RUN.md](DRY-RUN.md) for complete documentation on safe testing and multi-wing rollout.**

**Recommended workflow:**
1. Deploy to new wing
2. Run functions in dry-run mode (default)
3. Review logs in Application Insights
4. Set `EXECUTE=true` when confident
5. Run functions to apply changes

---

## Runtime Prerequisites

After deployment, the Azure Function App requires the following to operate:

**Azure Resources (created by Terraform):**
- Azure Function App with System-Assigned Managed Identity
- Azure Key Vault with CAPWATCH credentials stored
- Azure Storage Account for modules and data
- Application Insights for monitoring

**Microsoft Graph API Permissions (admin consent required):**
- `Directory.ReadWrite.All` - Manage directory objects
- `Group.ReadWrite.All` - Manage Microsoft 365 Groups
- `TeamMember.ReadWrite.All` - Manage Teams membership
- `User.Read.All` - Read user profiles
- `User.ReadWrite.All` - Create and update user accounts
- `Mail.Send` - Send email notifications (optional)

**Exchange Online API Permissions:**
- `Exchange.ManageAsApp` - Manage Exchange Online resources

**CAPWATCH Data Access:**
- Valid CAPWATCH credentials (stored in Key Vault)
- Network access to CAPWATCH API endpoints
- Scheduled download of CAPWATCH CSV data to Azure Storage

**Wing Configuration:**
- Environment variables configured in Function App settings (see [CONFIGURATION.md](CONFIGURATION.md))
- Wing-specific values for organizational structure

---

## Installation & Deployment

This toolkit is designed to be deployed as an Azure Function App using Infrastructure as Code (Terraform). The deployment process is streamlined for multi-wing adoption.

### 📋 Prerequisites

Before beginning deployment, ensure you have:

**Local Tools:**
- PowerShell 7.6+ installed on your local machine
- Azure CLI installed and authenticated (`az login`)
- Terraform installed (version >= 1.0)
- Azure Functions Core Tools (`func` command)
- Git for cloning the repository

**Azure Access:**
- Azure subscription with appropriate permissions
  - `Contributor` role on the subscription or resource group
  - `Application Administrator` role in Azure AD/Entra ID
- Ability to grant admin consent for Microsoft Graph API permissions

**Wing Information:**
- Your wing's 2-letter designator (e.g., CO, TX, CA, FL)
- CAPWATCH Organization ID (contact your wing IT staff if unknown)
- Exchange Online organization domain (e.g., `COCivilAirPatrol.onmicrosoft.com`)
- CAPWATCH username and password (same as eServices credentials)

### 🚀 Deployment Steps

Follow these steps in order for a successful deployment:

#### Step 1: Clone the Repository

```bash
git clone https://github.com/ezaero/CAPWATCHpwsh.git
cd CAPWATCHpwsh
```

#### Step 2: Configure Terraform Variables

```bash
cd terraform
cp terraform.tfvars.example terraform.tfvars
```

Edit `terraform.tfvars` with your wing's specific values:

```hcl
# Wing Configuration
wing_designator = "TX"  # Your wing's 2-letter code
location        = "East US"  # Azure region closest to your wing

# CAPWATCH Configuration
capwatch_org_id = "456"  # Your wing's CAPWATCH Organization ID
capwatch_username = "your-capwatch-username"  # eServices username
capwatch_password = "your-capwatch-password"  # eServices password

# Exchange Online Configuration
exchange_organization = "TXCivilAirPatrol.onmicrosoft.com"  # Your wing's domain

# Timezone (adjust for your location)
timezone = "Central Standard Time"

# Email Configuration (optional)
log_email_to_address   = "admin@yourwing.cap.gov"
log_email_from_address = "noreply@yourwing.cap.gov"
```

#### Step 3: Deploy Azure Infrastructure

```bash
# Initialize Terraform
terraform init

# Review the deployment plan
terraform plan

# Deploy the infrastructure
terraform apply
```

This creates:
- Resource Group
- Storage Account
- Function App (Consumption Plan)
- Key Vault (with CAPWATCH credentials)
- Application Insights
- Azure AD Application and Service Principal

**Important:** Note the output values after deployment, including:
- Resource Group name
- Storage Account name
- Function App name
- Key Vault name

#### Step 4: Grant Microsoft Graph API Permissions

After Terraform deployment, you must grant admin consent for Microsoft Graph API permissions:

1. Go to Azure Portal → Microsoft Entra ID → App registrations
2. Find your app (e.g., `CAPWATCHSync-TX`)
3. Go to API permissions
4. Click **Grant admin consent for [your tenant]**
5. Confirm all required permissions are consented

**Required Permissions:**
- `Directory.ReadWrite.All`
- `Group.ReadWrite.All`
- `TeamMember.ReadWrite.All`
- `User.Read.All`
- `User.ReadWrite.All`
- `Mail.Send`
- `Exchange.ManageAsApp` (Exchange Online)

#### Step 5: Set Up PowerShell Modules

Return to the project root directory and upload PowerShell modules to Azure Storage:

```powershell
# Return to project root
cd ..

# Connect to Azure
Connect-AzAccount

# Download required modules locally
./Download_Modules.ps1

# Upload modules to Azure Storage
# Replace with your actual storage account and resource group names from Terraform output
./Upload-ModulesToStorage.ps1 -StorageAccountName "capwatchtxsa" -ResourceGroup "capwatch-tx-rg"
```

This step is **critical** because Azure Functions has a 150MB deployment size limit. The required PowerShell modules (~180MB) are uploaded to Azure Storage and loaded at runtime.

#### Step 6: Deploy Function App Code

```bash
# Deploy the PowerShell functions to Azure
# Replace with your function app name from Terraform output
func azure functionapp publish capwatch-tx-func --powershell
```

#### Step 7: Verify Deployment

1. **Check Function App Status:**
   ```bash
   az functionapp show --name capwatch-tx-func --resource-group capwatch-tx-rg
   ```

2. **Test CAPWATCH Data Download:**
   - Go to Azure Portal → Function App → Functions
   - Find `download-extract-capwatch`
   - Click "Test/Run" to trigger manually
   - Check logs for successful execution

3. **Monitor Logs:**
   - Azure Portal → Function App → Application Insights
   - Review logs and traces for any errors

### 📖 Detailed Documentation

For more information on specific topics:

- **[DEPLOYMENT.md](DEPLOYMENT.md)** - Complete step-by-step deployment guide with troubleshooting
- **[MODULE-SETUP.md](MODULE-SETUP.md)** - PowerShell module management details
- **[CONFIGURATION.md](CONFIGURATION.md)** - Wing-specific configuration reference
- **[terraform/README.md](terraform/README.md)** - Terraform infrastructure documentation

### ⚡ Key Deployment Notes

- **Infrastructure as Code**: All Azure resources are defined in Terraform for reproducibility
- **Module Management**: PowerShell modules are uploaded to Azure Storage and loaded at runtime
- **Secure Credentials**: CAPWATCH credentials are stored in Azure Key Vault
- **Managed Identity**: Function App uses system-assigned managed identity for secure access
- **Multi-Wing Ready**: Configuration is parameterized for easy adaptation to any Wing

---

## Post-Deployment Configuration

After deploying the infrastructure and function code, complete these configuration steps:

### 1. Verify Function App Settings

Check that all required environment variables are set in Azure Portal:
- Navigate to Function App → Configuration → Application settings
- Verify: `WING_DESIGNATOR`, `CAPWATCH_ORGID`, `KEYVAULT_NAME`, `EXCHANGE_ORGANIZATION`

### 2. Test CAPWATCH Data Download

Manually trigger the data download function:
```bash
az functionapp function invoke --name capwatch-tx-func --function-name download-extract-capwatch --resource-group capwatch-tx-rg
```

Or use Azure Portal:
- Function App → Functions → `download-extract-capwatch` → Test/Run

Verify:
- Function executes successfully
- CSV files appear in `$HOME/data/CAPWatch`
- No authentication errors in logs

### 3. Configure Function Schedules

Review and adjust timer triggers in each function's `function.json`:
- `download-extract-capwatch`: Default daily at 2:00 AM
- `updateTeams`: Default daily at 3:00 AM
- `checkAccounts`: Default daily at 3:30 AM
- `DLOpsQuals`: Default daily at 4:00 AM
- `Maintenance`: Default monthly on the 1st at 5:00 AM

### 4. Enable Dry-Run Mode Initially

For functions that make changes, enable dry-run mode for initial testing:
- Set `EXECUTE=false` in Function App settings
- Monitor logs to verify expected behavior
- When satisfied, set `EXECUTE=true` to enable actual changes

### 5. Monitor and Validate

- Check Application Insights for errors or warnings
- Review logs in `$HOME/logs` directory
- Verify Teams, users, and distribution groups are being updated correctly
- Monitor for any permission or authentication issues

---

## Usage

### Automated Execution

Once deployed and configured, the functions run automatically on their schedules:
- **No manual intervention required** for daily operations
- All logs written to `$HOME/logs` directory
- Email notifications sent to configured addresses (if enabled)

### Manual Execution

You can manually trigger any function:

**Via Azure Portal:**
- Function App → Functions → Select function → Code + Test → Test/Run

**Via Azure CLI:**
```bash
az functionapp function invoke --name <function-app-name> --function-name <function-name> --resource-group <resource-group>
```

### Common Functions

- **download-extract-capwatch**: Download latest CAPWATCH data
- **updateTeams**: Synchronize Teams membership with CAPWATCH
- **checkAccounts**: Create/update user accounts and mail contacts
- **DLOpsQuals**: Manage operational qualifications distribution groups
- **DLAnnouncements**: Manage announcement distribution lists
- **Maintenance**: Monthly cleanup of expired accounts and old logs

### Advanced Configuration

**Dry-Run Mode**: Many functions support dry-run mode for testing
- Set `EXECUTE=false` in Function App settings to preview changes without applying them
- Review logs to verify expected behavior
- Set `EXECUTE=true` when ready to enable live updates

**Custom Team Mappings** (`updateTeams`):
- Synchronizes distribution list membership with Teams
- Configure via `group_to_team.csv` file or `GROUP_TEAM_PAIRS` environment variable
- Automatically derives distribution list addresses from Team names
- See [updateTeams/readme.md](updateTeams/readme.md) for detailed configuration

**Force Mode**:
- Set `FORCE=true` to skip interactive confirmations (required for automated execution)
- Always use dry-run mode first to validate changes


---

## Project Structure

### Azure Functions (Timer-Triggered)

- **`/download-extract-capwatch`** – Downloads and extracts CAPWATCH CSV data files from CAPWATCH API (Daily: 2:00 AM)
- **`/updateTeams`** – Synchronizes Microsoft Teams membership and ownership with CAPWATCH unit data (Daily: 3:00 AM)
- **`/checkAccounts`** – Creates, updates, and restores user accounts and mail contacts in Azure AD and Exchange (Daily: 3:30 AM)
- **`/update-user-names-ranks`** – Updates user display names and attributes to reflect current rank and name from CAPWATCH (Daily: 3:45 AM)
- **`/DLAnnouncements`** – Manages announcement distribution lists with membership from CAPWATCH (Daily: 4:00 AM)
- **`/DLOpsQuals`** – Automates operational qualifications distribution groups (pilots, aircrew, ES) (Daily: 4:15 AM)
- **`/DLSeniorsCadets`** – Maintains distribution lists for senior and cadet members (Daily: 4:30 AM)
- **`/DLSpecTrack`** – Manages specialty track distribution lists for targeted communications (Daily: 4:45 AM)
- **`/Maintenance`** – Performs monthly cleanup: deletes expired accounts and old log files (Monthly: 1st at 5:00 AM)
- **`/emailLogFile`** – Sends log file summaries via email to administrators (Daily: 6:00 AM)

### Operational Flight Scheduling

- **`/OFlights`** – Manages operational flight scheduling and tracking
- **`/OFlightMetrics`** – Calculates and reports operational flight metrics
- **`/escalatePilotInvitations`** – Handles pilot invitation reminders and escalations
- **`/sendReminders`** – Sends automated reminders for operational flights
- **`/processInvitationExpiry`** – Processes expired pilot invitations

### Support and Utilities

- **`/shared`** – Shared utility functions, logging, and Microsoft Graph authentication helpers
  - `shared.ps1` – Common functions used across all functions
  - `Load-Modules.ps1` – Runtime PowerShell module loading from Azure Storage
- **`/terraform`** – Infrastructure as Code (Terraform) for Azure resource deployment
- **`Download_Modules.ps1`** – Downloads required PowerShell modules locally for upload
- **`Upload-ModulesToStorage.ps1`** – Uploads PowerShell modules to Azure Storage
- **`profile.ps1`** – Function App startup configuration and module initialization
- **`requirements.psd1`** – Azure Functions managed dependencies configuration
- **`host.json`** – Azure Functions host configuration
- **`.funcignore`** – Deployment exclusions to optimize package size

### Documentation

- **`README.md`** – This file: overview and quick start guide
- **`DEPLOYMENT.md`** – Detailed step-by-step deployment instructions
- **`CONFIGURATION.md`** – Wing-specific configuration reference
- **`MODULE-SETUP.md`** – PowerShell module management quick reference
- **`terraform/README.md`** – Terraform infrastructure documentation

---

## Troubleshooting

### Common Issues and Solutions

**Module Loading Failures**
- **Symptom**: Functions fail with "module not found" errors
- **Solution**:
  - Verify modules uploaded to Azure Storage: Portal → Storage Account → modules container
  - Check Function App has Storage Blob Data Reader permissions
  - Review `shared/Load-Modules.ps1` logs for download failures
  - Manually trigger module re-upload: `./Upload-ModulesToStorage.ps1`

**CAPWATCH Authentication Failures**
- **Symptom**: `download-extract-capwatch` fails with 401 or 403 errors
- **Solution**:
  - Verify credentials in Key Vault are correct (same as eServices login)
  - Check Function App has Get/List permissions on Key Vault
  - Verify CAPWATCH Organization ID is correct for your Wing
  - Test credentials manually at eServices.cap.gov

**Microsoft Graph Permission Errors**
- **Symptom**: Functions fail with "Insufficient permissions" errors
- **Solution**:
  - Verify admin consent granted: Portal → Entra ID → App registrations → API permissions
  - Check all required permissions are listed and consented
  - Ensure Function App managed identity is enabled
  - Wait 5-10 minutes after granting consent for propagation

**Stale CAPWATCH Data**
- **Symptom**: Functions stop with "CAPWATCH data is stale" message
- **Solution**:
  - Check `download-extract-capwatch` function is running successfully
  - Verify network connectivity to CAPWATCH API
  - Review Application Insights logs for download failures
  - Manually trigger download function to refresh data

**Terraform Deployment Failures**
- **Symptom**: `terraform apply` fails with permission or quota errors
- **Solution**:
  - Verify Azure subscription has sufficient quota for resources
  - Check you have Contributor role on subscription/resource group
  - Ensure you have Application Administrator role in Azure AD
  - Review Terraform error messages for specific resource failures

**Function Deployment Size Exceeded**
- **Symptom**: `func publish` fails with "deployment package exceeds size limit"
- **Solution**:
  - Verify `.funcignore` excludes Modules/, terraform/, logs/, data/ directories
  - Check modules were uploaded to Azure Storage (not included in deployment)
  - Review deployment package contents: `func pack --build-native-deps`
  - Expected size: ~53KB after exclusions

### Getting Support

1. **Check Logs**: Review Application Insights and function logs in `$HOME/logs`
2. **Review Documentation**: Consult DEPLOYMENT.md and CONFIGURATION.md
3. **GitHub Issues**: Report bugs or request features at repository issues page
4. **Wing IT Staff**: Contact your wing's IT department for organization-specific questions

---

## Security & Best Practices

- **Do not commit secrets or credentials.** Use environment variables or Azure Key Vault for sensitive data.
- **Review all scripts for organization-specific information** before making the repository public.
- **Follow [Azure best practices](https://learn.microsoft.com/azure/cloud-adoption-framework/ready/azure-best-practices/)** for automation and security.
- **Use a `.gitignore`** to exclude logs, output, credentials, and IDE files.

---

## Azure Cost Estimation

Running this solution on Azure incurs costs. Typical monthly costs for a Wing:

### Expected Costs (Approximate)

| Resource | Tier | Estimated Monthly Cost |
|----------|------|------------------------|
| **Function App** | Consumption Plan | $5-15 (based on executions) |
| **Storage Account** | Standard LRS | $1-3 (minimal data storage) |
| **Key Vault** | Standard | $0-1 (few secrets) |
| **Application Insights** | Standard | $5-10 (based on data ingestion) |
| **Total** | | **$11-29/month** |

### Cost Optimization Tips

- **Consumption Plan**: Function App only charges for execution time (no idle costs)
- **Storage**: Use Standard LRS (not Premium) for cost efficiency
- **Application Insights**: Configure data sampling to reduce ingestion costs
- **Scheduled Functions**: Optimize timer triggers to avoid unnecessary executions
- **Logs**: Implement log retention policies to control storage costs

### Cost Factors

Actual costs vary based on:
- Wing size (number of members and units)
- Execution frequency (daily vs. hourly schedules)
- Application Insights data ingestion volume
- Storage usage for logs and CAPWATCH data
- Number of Microsoft Graph API calls

**Note**: These costs do not include Microsoft 365 licensing, which is required separately for user accounts.

---

## License

This project is licensed under the [MIT License](LICENSE).

**Copyright (c) 2024 Civil Air Patrol Wings**

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED.

---

## Frequently Asked Questions

### General Questions

**Q: Does this work with my Wing?**
A: Yes! This solution is designed to work with any CAP Wing. You just need to configure it with your Wing's specific values (designator, CAPWATCH Org ID, Exchange domain).

**Q: How long does deployment take?**
A: Initial deployment takes 2-3 hours including infrastructure setup, module upload, and testing. Subsequent updates take minutes.

**Q: What if my Wing has unique requirements?**
A: The solution is highly configurable. Most Wings can use it as-is, but you can customize functions for Wing-specific needs while maintaining the core functionality.

**Q: Do I need to be a PowerShell expert?**
A: No. The deployment uses Terraform and automated scripts. Basic PowerShell knowledge is helpful for troubleshooting, but detailed deployment documentation is provided.

### Technical Questions

**Q: Why use Azure Functions instead of a VM?**
A: Azure Functions are serverless (no OS patching), automatically scaling, and cost-effective (pay only for execution time). A VM would require ongoing maintenance and cost more.

**Q: What happens if CAPWATCH is down?**
A: Functions will log errors and retry on the next scheduled execution. No changes are made if CAPWATCH data cannot be retrieved.

**Q: Can I run this on-premises instead of Azure?**
A: The solution is designed for Azure, but you could adapt it to run on-premises with modifications. However, you would lose the serverless benefits and need to manage infrastructure yourself.

**Q: How are credentials protected?**
A: CAPWATCH credentials are stored in Azure Key Vault with encryption at rest. The Function App uses managed identity (no credentials in code). All API access uses OAuth tokens.

**Q: What if I make a mistake during deployment?**
A: Terraform makes it easy to destroy and redeploy resources. During initial testing, use dry-run mode (`EXECUTE=false`) to preview changes before enabling live updates.

### Cost and Licensing Questions

**Q: How much does this cost to run?**
A: Typically $11-29/month for Azure resources (Function App, Storage, Key Vault, Application Insights). See [Azure Cost Estimation](#azure-cost-estimation) for details.

**Q: Do I need separate Microsoft 365 licenses?**
A: Yes, users need Microsoft 365 licenses (typically assigned by CAP National or Wing). This solution manages accounts but doesn't provide licenses.

**Q: Are there any hidden costs?**
A: No hidden costs, but be aware of: (1) Microsoft Graph API calls are free but rate-limited, (2) Application Insights ingestion scales with usage, (3) Storage grows with logs (implement retention).

### Operational Questions

**Q: How often does data sync?**
A: By default, CAPWATCH data downloads daily and synchronization runs shortly after. You can adjust schedules in each function's `function.json`.

**Q: What happens to accounts when members leave?**
A: User accounts are soft-deleted in Azure AD. If a member rejoins within 30 days, their account is automatically restored.

**Q: Can I run functions manually?**
A: Yes! You can manually trigger any function via Azure Portal or Azure CLI for testing or immediate updates.

**Q: How do I know if something goes wrong?**
A: Check Application Insights for errors, review function logs in `$HOME/logs`, and optionally configure email notifications for daily log summaries.

### Security and Compliance Questions

**Q: Is this secure enough for CAP data?**
A: Yes. The solution follows Azure security best practices: managed identities, Key Vault for secrets, encrypted storage, least-privilege permissions, and audit logging.

**Q: Who has access to the data?**
A: Only the Function App's managed identity and authorized Azure administrators. Access is controlled through Azure RBAC and Key Vault access policies.

**Q: Does this comply with CAP regulations?**
A: The solution is a technical implementation. Your Wing is responsible for ensuring deployment and usage comply with CAP policies and data protection requirements.

**Q: Can I audit what changes were made?**
A: Yes. All operations are logged with timestamps and details. Application Insights provides searchable telemetry, and logs include before/after states.

---

## Multi-Wing Collaboration

### Wings Using This Solution

This solution is actively deployed and maintained across multiple CAP Wings:
- **Colorado Wing (CO)** - Original implementation and primary development
- **Additional Wings** - Expanding to other Wings through collaborative deployment

### Sharing Improvements

If your Wing develops enhancements, bug fixes, or improvements:
1. Test thoroughly in your environment
2. Document changes and configuration requirements
3. Submit pull requests to benefit all Wings
4. Share lessons learned and best practices

### Wing-Specific Customizations

Each Wing may need customizations for:
- Local organizational structure differences
- Specialty distribution lists unique to your Wing
- Custom reporting requirements
- Integration with Wing-specific systems

Document these customizations separately and avoid committing Wing-specific data to the shared repository.

---

## Contributing

Contributions from all Wings are welcome and encouraged:

- **Bug Reports**: Open an issue with detailed reproduction steps
- **Feature Requests**: Describe the use case and benefit to multiple Wings
- **Pull Requests**: Follow existing code structure and include documentation
- **Documentation**: Improvements to deployment guides, troubleshooting, and configuration

Before contributing:
1. Ensure changes work across different Wing configurations
2. Remove any Wing-specific credentials or sensitive data
3. Test with multiple organizational structures if applicable
4. Update relevant documentation

---

## Support and Contact

### For Deployment Questions
- Review **[DEPLOYMENT.md](DEPLOYMENT.md)** for detailed instructions
- Check **[Troubleshooting](#troubleshooting)** section above
- Consult your Wing's IT staff for organization-specific configuration

### For Technical Issues
- **GitHub Issues**: Report bugs and technical problems
- **Application Insights**: Check function logs for error details
- **Azure Support**: For Azure infrastructure or platform issues

### For Wing-Specific Configuration
- Contact your Wing's IT department
- Review **[CONFIGURATION.md](CONFIGURATION.md)** for required values
- Verify CAPWATCH Organization ID with Wing leadership

---

## Disclaimer

This project is provided as-is and is not officially supported or endorsed by Civil Air Patrol National Headquarters or Microsoft Corporation.

- Use at your own risk and responsibility
- Test thoroughly before deploying to production
- Maintain appropriate backups and disaster recovery procedures
- Ensure compliance with CAP policies and regulations
- Verify all changes align with your Wing's IT governance

Each Wing is responsible for:
- Securing their Azure environment and credentials
- Maintaining appropriate access controls
- Monitoring costs and resource usage
- Ensuring data privacy and security compliance
