# Terraform Infrastructure for CAPWATCHSyncPWSH

This directory contains Terraform configuration files to provision all Azure resources needed to run the CAPWATCH synchronization functions.

## 📋 Prerequisites

1. **Azure CLI** installed and authenticated (`az login`)
2. **Terraform** installed (version >= 1.0)
3. **Azure subscription** with appropriate permissions to create resources
4. **CAPWATCH credentials** (username and password for CAPWATCH API)

## 🚀 Quick Start

1. **Clone and navigate to terraform directory:**
   ```bash
   cd terraform
   ```

2. **Create your configuration file:**
   ```bash
   cp terraform.tfvars.example terraform.tfvars
   ```


3. **Edit `terraform.tfvars` with your wing's specific values:**
   ```hcl
   wing_designator = "TX"  # Your wing's 2-letter code
   capwatch_org_id = "456" # Your wing's CAPWATCH Organization ID
   exchange_organization = "TXCivilAirPatrol.onmicrosoft.com"
   ```

4. **After deployment, add CAPWATCH credentials to Key Vault:**
   ```bash
   az keyvault secret set --vault-name <your-keyvault-name> --name capwatch-username --value "your-capwatch-username"
   az keyvault secret set --vault-name <your-keyvault-name> --name capwatch-password --value "your-capwatch-password"
   ```

4. **Initialize Terraform:**
   ```bash
   terraform init
   ```

5. **Plan the deployment:**
   ```bash
   terraform plan
   ```

6. **Deploy the infrastructure:**
   ```bash
   terraform apply
   ```

## 📁 File Structure

- `main.tf` - Main infrastructure resources
- `variables.tf` - Input variables and validation
- `outputs.tf` - Output values and deployment instructions
- `secrets.tf` - Key Vault secrets management
- `terraform.tfvars.example` - Example configuration file

## 🏗️ Resources Created

| Resource | Purpose |
|----------|---------|
| **Resource Group** | Container for all resources |
| **Function App** | Hosts the PowerShell automation functions |
| **App Service Plan** | Consumption plan for serverless execution |
| **Storage Account** | Required for Function App operation |
| **Key Vault** | Securely stores CAPWATCH credentials |
| **Application Insights** | Monitoring and logging |
| **Azure AD Application** | Enterprise app for Microsoft Graph permissions |
| **Service Principal** | Identity for API access |

## ⚙️ Configuration Variables

### Required Variables

| Variable | Description | Example |
|----------|-------------|---------|
| `wing_designator` | Two-letter wing code | `"TX"` |
| `capwatch_org_id` | CAPWATCH Organization ID | `"456"` |
| `exchange_organization` | Exchange Online domain | `"TXCivilAirPatrol.onmicrosoft.com"` |
| `capwatch_username` | CAPWATCH username | `"john.doe"` |
| `capwatch_password` | CAPWATCH password | `"secretpassword"` |

### Optional Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `location` | Azure region | `"East US"` |
| `timezone` | Function App timezone | `"Mountain Standard Time"` |
| `log_email_to_address` | Email for notifications | `""` |
| `log_email_from_address` | Sender email | `""` |

## 🔐 Security Features

- **Managed Identity** for secure Azure resource access
- **Key Vault** for credential storage with proper access policies
- **Least-privilege** permissions on all resources
- **No plaintext secrets** in configuration files

## 📋 Post-Deployment Steps

After successful deployment, follow these steps:


1. **Grant Microsoft Graph API permissions in Microsoft Entra ID:**
   - Go to Azure Portal > Microsoft Entra ID > **Manage** > **App registrations** (for application permissions)
   - (You can also use **Enterprise applications** to view the service principal, but admin consent is granted in App registrations)
   - Find your app (e.g., CAPWATCHSync-CO)
   - Go to **API permissions**
   - Click **Grant admin consent for [your tenant]**
   - Confirm that the required Microsoft Graph permissions are listed and consented

2. **Deploy Function App code** using Azure Functions Core Tools
3. **Test the functions** starting with `download-extract-capwatch`
4. **Monitor logs** in Application Insights

## 🧪 Testing

1. **Validate deployment:**
   ```bash
   terraform validate
   ```

2. **Check resource status:**
   ```bash
   az resource list --resource-group $(terraform output -raw resource_group_name)
   ```

3. **Test Key Vault access:**
   ```bash
   az keyvault secret show --vault-name $(terraform output -raw key_vault_name) --name capwatch-username
   ```

## 🗑️ Cleanup

To remove all resources:
```bash
terraform destroy
```

## 🆘 Troubleshooting

### Common Issues

1. **Insufficient permissions:**
   - Ensure you have `Contributor` role on the subscription
   - Ensure you have `Application Administrator` role in Azure AD

2. **Key Vault access denied:**
   - Verify your user has appropriate Key Vault access policies
   - Check that the Function App managed identity is granted access

3. **Function App deployment fails:**
   - Verify all environment variables are set correctly
   - Check Application Insights logs for detailed error messages

### Getting Help

1. Check Terraform plan output before applying
2. Review Azure Activity Logs for resource creation issues
3. Use `terraform refresh` to sync state with actual resources
4. Enable Terraform debug logging: `export TF_LOG=DEBUG`

## 📝 Notes

- The Azure AD application requires **admin consent** for Microsoft Graph permissions
- CAPWATCH credentials are encrypted at rest in Key Vault
- Function App uses **system-assigned managed identity** for secure access
- All resources are tagged for easy management and cost tracking
