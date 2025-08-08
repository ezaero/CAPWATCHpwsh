# Configuration Guide for Multi-Wing Deployment

This guide helps you configure CAPWATCHSyncPWSH for any CAP Wing.

## Required Environment Variables

Configure these environment variables in your Azure Function App settings:

### Wing-Specific Configuration

| Variable | Description | Example |
|----------|-------------|---------|
| `WING_DESIGNATOR` | Two-letter wing abbreviation | `CO`, `TX`, `CA`, etc. |
| `CAPWATCH_ORGID` | Your wing's CAPWATCH Organization ID | `123` (check with your wing IT staff) |
| `KEYVAULT_NAME` | Name of your Azure Key Vault | `yourwing-capwatch-kv` |
| `EXCHANGE_ORGANIZATION` | Your wing's Exchange Online organization | `YourWingCivilAirPatrol.onmicrosoft.com` |

### Email Configuration (Optional)

| Variable | Description | Example |
|----------|-------------|---------|
| `LOG_EMAIL_TO_ADDRESS` | Email address to receive log notifications | `admin@yourwing.cap.gov` |
| `LOG_EMAIL_FROM_ADDRESS` | Sender email address for notifications | `noreply@yourwing.cap.gov` |

## Azure Resources Setup

### 1. Azure Key Vault

Create an Azure Key Vault with the name specified in `KEYVAULT_NAME` and add these secrets:

- `capwatch-username`: Your CAPWATCH username (same as eServices login)
- `capwatch-password`: Your CAPWATCH password (same as eServices login)

Grant your Azure Function App's managed identity **Get** and **List** permissions to this Key Vault.

### 2. Exchange Online Organization

Your `EXCHANGE_ORGANIZATION` should match your wing's Exchange Online tenant. This is typically in the format:
`[WingName]CivilAirPatrol.onmicrosoft.com`

Examples:
- Colorado: `COCivilAirPatrol.onmicrosoft.com`
- Texas: `TXCivilAirPatrol.onmicrosoft.com`
- California: `CACivilAirPatrol.onmicrosoft.com`

### 3. Microsoft Graph API Permissions

Ensure your Azure Function App's managed identity has these Microsoft Graph API permissions:

- `Group.ReadWrite.All`
- `TeamMember.ReadWrite.All`
- `User.Read.All`
- `User.ReadWrite.All`
- `Directory.ReadWrite.All`
- `Mail.Send` (if using email notifications)

## Wing-Specific Customization

### Finding Your CAPWATCH Organization ID

1. Contact your wing's IT staff or CAPWATCH administrator
2. The Organization ID is typically a 3-4 digit number unique to your wing
3. This ID is used in the CAPWATCH API to download your wing's data

### Team and User Naming Conventions

The scripts will automatically use your `WING_DESIGNATOR` to:
- Create team names like `[WING_DESIGNATOR]-[3-digit UnitNumber]` (e.g., `TX-001`, `CA-075`)
- Generate user principal names with your Exchange organization domain
- Log activities with wing-specific identifiers

## Local Development

1. Copy `local.settings.example.json` to `local.settings.json`
2. Update the values with your wing's specific configuration
3. Test with a subset of units before deploying to production

## Deployment Checklist

- [ ] Azure Key Vault created with CAPWATCH credentials
- [ ] Function App managed identity configured
- [ ] Microsoft Graph API permissions granted
- [ ] Environment variables set in Function App configuration
- [ ] Exchange Online organization verified
- [ ] CAPWATCH Organization ID confirmed
- [ ] Test run completed with sample data

## Support

For questions about wing-specific configuration:
1. Consult your wing's IT staff for organization-specific values
2. Review Azure Function App logs for configuration issues
3. Verify Key Vault access and permissions if authentication fails
