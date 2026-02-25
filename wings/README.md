# Wings Configuration - Local Settings

Each wing folder contains local configuration files used for deployment and development.

## Structure

- `colorado/` - Colorado Wing configuration
- `montana/` - Montana Wing configuration

## Files in Each Wing Folder

### `local.settings.json`
Local development configuration for the Azure Functions runtime. Contains wing-specific environment variables.

**⚠️ Security Note**: Never commit `local.settings.json` files with real credentials to Git.

### `terraform.tfvars`
Terraform variables for provisioning cloud infrastructure specific to the wing.

**⚠️ Security Note**: Be careful not to commit sensitive data like API keys or credentials in this file.

## Usage

### Local Development

Copy the appropriate wing's settings file to the root directory:

```bash
# For Colorado Wing development
cp wings/colorado/local.settings.json ./local.settings.json

# For Montana Wing development
cp wings/montana/local.settings.json ./local.settings.json
```

Then update `AzureWebJobsStorage` with your Azure Storage connection string.

### Deployment

When deploying infrastructure, specify the wing's terraform variables:

```bash
# Deploy Colorado infrastructure
terraform apply -var-file="../wings/colorado/terraform.tfvars"

# Deploy Montana infrastructure
terraform apply -var-file="../wings/montana/terraform.tfvars"
```

## Adding a New Wing

1. Create a new directory: `wings/{wing_abbreviation}/`
2. Copy `local.settings.json` from an existing wing
3. Update `WING_DESIGNATOR` and other wing-specific values
4. Copy `terraform.tfvars` from an existing wing
5. Update all wing-specific variables
6. Commit the new wing configuration to Git

See [WING-DEPLOYMENT.md](WING-DEPLOYMENT.md) for detailed steps.

## Environment Variables by Wing

| Variable | Colorado | Montana |
|----------|----------|---------|
| `WING_DESIGNATOR` | CO | MT |
| `CAPWATCH_ORGID` | 123 | 456 |
| `EXCHANGE_ORGANIZATION` | COCivilAirPatrol.onmicrosoft.com | MTCivilAirPatrol.onmicrosoft.com |
| `KEYVAULT_NAME` | cowg-capwatch-kv | mtwg-capwatch-kv |
| `COSMOS_DATABASE` | orientation-flights | orientation-flights-mt |

---

For detailed multi-wing deployment information, see [WING-DEPLOYMENT.md](WING-DEPLOYMENT.md).
