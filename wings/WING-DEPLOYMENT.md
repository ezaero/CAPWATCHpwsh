# Multi-Wing Deployment Guide

This repository supports deployment to multiple CAP Wings using a configuration-driven approach. The shared codebase ensures all wings benefit from improvements and bug fixes, while allowing wing-specific customization.

## Repository Structure

```
/
├── shared/                          # Shared modules used by all wings
│   ├── shared.ps1                   # Common PowerShell functions
│   └── Load-Modules.ps1             # Module loading utilities
├── wings/                           # Wing-specific configurations
│   ├── colorado/                    # Colorado Wing
│   │   ├── local.settings.json      # Colorado local development config
│   │   └── terraform.tfvars         # Colorado Azure infrastructure config
│   └── montana/                     # Montana Wing
│       ├── local.settings.json      # Montana local development config
│       └── terraform.tfvars         # Montana Azure infrastructure config
├── [function folders]/              # Shared Azure Functions
│   ├── escalatePilotInvitations/
│   ├── sendReminders/
│   ├── updateTeams/
│   └── ... (all other functions)
├── terraform/                       # Infrastructure as Code templates
├── CONFIGURATION.md                 # General configuration guide
└── DEPLOYMENT.md                    # Deployment instructions
```

## Key Concepts

### Configuration-Driven Design

All wing-specific settings are externalized to configuration:

- **Local Development**: Use the wing-specific `local.settings.json` file
- **Azure Deployment**: Environment variables are set via Azure Functions app settings
- **Infrastructure**: Terraform variables defined in wing-specific `terraform.tfvars`

### Wing Designators

Each wing has a 2-letter code used throughout the system:
- `CO` = Colorado Wing
- `MT` = Montana Wing
- `TX` = Texas Wing (future), etc.

## Deploying Colorado Wing

### Prerequisites

1. **Azure Subscription** with appropriate permissions
2. **Terraform** v1.0+
3. **Azure CLI** installed and authenticated
4. **PowerShell 7.6+** for local testing
5. **Git** for version control

### Step 1: Local Development Setup

```bash
# 1. Clone the repository
git clone https://github.com/ezaero/CAPWATCHpwsh.git
cd CAPWATCHpwsh

# 2. Copy Colorado's local settings
cp wings/colorado/local.settings.json ./local.settings.json

# 3. Update local.settings.json with your Azure Storage Connection String
# Find: "AzureWebJobsStorage": ""
# Replace with your actual connection string
```

### Step 2: Azure Infrastructure Setup

```bash
# 1. Navigate to terraform directory
cd terraform

# 2. Initialize terraform
terraform init

# 3. Plan the infrastructure (Colorado)
terraform plan -var-file="../wings/colorado/terraform.tfvars"

# 4. Apply the configuration
terraform apply -var-file="../wings/colorado/terraform.tfvars"
```

**Output**: 
- Azure Function App created
- Cosmos DB instance configured
- Application Insights enabled
- Key Vault configured

### Step 3: Configure Azure Key Vault

From the Azure Portal or Azure CLI:

```bash
# 1. Get your Key Vault name
KEYVAULT_NAME="cowg-capwatch-kv"

# 2. Add CAPWATCH credentials
az keyvault secret set --vault-name $KEYVAULT_NAME --name "capwatch-username" --value "your-username"
az keyvault secret set --vault-name $KEYVAULT_NAME --name "capwatch-password" --value "your-password"

# 3. Add Cosmos DB connection string
az keyvault secret set --vault-name $KEYVAULT_NAME --name "cosmos-connection-string" --value "your-connection-string"
```

### Step 4: Deploy Functions

```bash
# From the repository root
func azure functionapp publish cowg-capwatch-app --build remote
```

### Step 5: Verify Deployment

1. Check Azure Portal for Function App status
2. Review Application Insights for logs
3. Test one function manually from the portal

---

## Deploying Montana Wing

The process is identical to Colorado, but uses Montana-specific configurations:

### Step 1: Local Development Setup

```bash
# Copy Montana's local settings
cp wings/montana/local.settings.json ./local.settings.json

# Update with your Azure Storage Connection String
```

### Step 2: Azure Infrastructure Setup

```bash
cd terraform

terraform init

# Use Montana's terraform variables
terraform plan -var-file="../wings/montana/terraform.tfvars"
terraform apply -var-file="../wings/montana/terraform.tfvars"
```

### Step 3: Configure Key Vault (Montana)

```bash
KEYVAULT_NAME="mtwg-capwatch-kv"

az keyvault secret set --vault-name $KEYVAULT_NAME --name "capwatch-username" --value "your-username"
az keyvault secret set --vault-name $KEYVAULT_NAME --name "capwatch-password" --value "your-password"
az keyvault secret set --vault-name $KEYVAULT_NAME --name "cosmos-connection-string" --value "your-connection-string"
```

### Step 4: Deploy Functions

```bash
func azure functionapp publish mtwg-capwatch-app --build remote
```

---

## Adding a New Wing

To support a new wing (e.g., Texas Wing):

### 1. Create Wing Configuration Folder

```bash
mkdir -p wings/texas
```

### 2. Create Configuration Files

Create `wings/texas/local.settings.json`:
```json
{
  "IsEncrypted": false,
  "Values": {
    "AzureWebJobsStorage": "",
    "FUNCTIONS_WORKER_RUNTIME": "powershell",
    "FUNCTIONS_EXTENSION_VERSION": "~4",
    
    "WING_DESIGNATOR": "TX",
    "CAPWATCH_ORGID": "789",
    "KEYVAULT_NAME": "txwg-capwatch-kv",
    "EXCHANGE_ORGANIZATION": "TXCivilAirPatrol.onmicrosoft.com",
    "COSMOS_CONNECTION_STRING": "",
    "COSMOS_DATABASE": "orientation-flights-tx",
    
    "LOG_EMAIL_TO_ADDRESS": "admin@txwg.cap.gov",
    "LOG_EMAIL_FROM_ADDRESS": "noreply@txwg.cap.gov",
    "FRONTEND_URL": "https://orientationflights.txwg.cap.gov",
    
    "WEBSITE_TIME_ZONE": "Mountain Standard Time",
    "PSWorkerInProcConcurrencyUpperBound": "1"
  }
}
```

Create `wings/texas/terraform.tfvars`:
```hcl
wing_designator          = "TX"
location                 = "Central US"
capwatch_org_id          = "789"
exchange_organization    = "TXCivilAirPatrol.onmicrosoft.com"
timezone                 = "Central Standard Time"
log_email_to_address     = "admin@txwg.cap.gov"
log_email_from_address   = "noreply@txwg.cap.gov"
```

### 3. Follow Deployment Steps

Use the same process as Colorado/Montana but with Texas-specific values.

---

## Managing Environment Variables

### Local Development

Edit `local.settings.json` in the wings/{wing}/ folder:

```json
{
  "Values": {
    "WING_DESIGNATOR": "CO",
    "CAPWATCH_ORGID": "123",
    // ... other settings
  }
}
```

### Azure Function App

Set via Azure Portal or Azure CLI:

```bash
az functionapp config appsettings set \
  --name cowg-capwatch-app \
  --resource-group cowg-capwatch-rg \
  --settings "WING_DESIGNATOR=CO" "CAPWATCH_ORGID=123"
```

Or use Terraform to manage all settings automatically.

---

## Wing-Specific Data Files

Some data files may be wing-specific (e.g., CSVs with unit lists). Consider organizing them:

```
/data
  /colorado
    units.csv
    cadets.csv
  /montana
    units.csv
    cadets.csv
```

Update function scripts to reference the correct wing-specific data:
```powershell
$dataFile = "$PSScriptRoot\..\data\$($env:WING_DESIGNATOR.ToLower())\units.csv"
```

---

## Updating Shared Code

When you update shared functions or modules:

1. **Test thoroughly** against all wing configurations
2. **Update BOTH wings** in Azure at deployment time
3. **Use feature flags** for gradual rollouts:

```powershell
if ($env:WING_DESIGNATOR -eq "CO") {
    # Colorado-specific implementation
} else {
    # Montana-specific implementation
}
```

---

## Troubleshooting Multi-Wing Deployments

### Issue: Wrong configuration applied

**Solution**: Verify you're using the correct terraform.tfvars file
```bash
terraform plan -var-file="../wings/{wing}/terraform.tfvars"
```

### Issue: Cosmos DB connection failing

**Solution**: Verify the connection string in Key Vault
```bash
az keyvault secret show --vault-name {keyvault-name} --name "cosmos-connection-string"
```

### Issue: Functions deployed but not running

**Solution**: Check Application Insights logs for the specific wing
```bash
# View last 100 errors for Colorado wing
az monitor app-insights query \
  --app cowg-capwatch-insights \
  --analytics-query "exceptions | where timestamp > ago(1d) | project timestamp, message | take 100"
```

---

## Best Practices

### ✅ Do's

- Keep shared code in `/shared` and function directories
- Externalize all wing-specific settings to configuration
- Use the same deployment script for all wings
- Test configuration changes against a test wing first
- Document any wing-specific customizations

### ❌ Don'ts

- Hard-code wing designators or organization IDs in functions
- Create separate code paths for each wing if avoidable
- Store sensitive data in version control
- Deploy without testing both wings

---

## Support and Questions

For wing-specific deployment issues, consult:
1. **[CONFIGURATION.md](../CONFIGURATION.md)** - General configuration details
2. **[DEPLOYMENT.md](../DEPLOYMENT.md)** - Deployment processes
3. Your wing's IT staff for organization-specific values
4. Azure Portal logs and Application Insights for runtime issues

---

## Version History

| Date | Wing | Action |
|------|------|--------|
| 2026-02-25 | Colorado | Initial deployment |
| 2026-02-25 | Montana | Added to multi-wing structure |
