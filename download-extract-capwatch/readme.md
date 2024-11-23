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
