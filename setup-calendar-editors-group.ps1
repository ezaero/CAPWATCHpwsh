# Setup script for COWG-calendar-editors mail-enabled security group
# This script creates the group and configures permissions for the managed identity
# Run this script with Exchange Online admin credentials before running the sync function

param(
    [Parameter(Mandatory=$false)]
    [string]$ManagedIdentityObjectId,
    
    [Parameter(Mandatory=$false)]
    [string]$GroupName = "COWG-calendar-editors"
)

Write-Host "Setting up mail-enabled security group: $GroupName" -ForegroundColor Cyan

# If ManagedIdentityObjectId not provided, try to get it from the function app
if (-not $ManagedIdentityObjectId) {
    Write-Host "Attempting to retrieve managed identity object ID from Azure resources..." -ForegroundColor Yellow
    
    # Get current Az context
    $context = Get-AzContext
    if (-not $context) {
        Write-Error "Not authenticated to Azure. Please run 'Connect-AzAccount' first."
        exit 1
    }
    
    # Try to get the function app managed identity
    try {
        # Get the function app - adjust the resource group name as needed
        $functionApp = Get-AzFunctionApp -ErrorAction Stop | Where-Object { $_.Runtime -like "*PowerShell*" } | Select-Object -First 1
        
        if ($functionApp) {
            $ManagedIdentityObjectId = $functionApp.IdentityPrincipalId
            Write-Host "Found managed identity: $ManagedIdentityObjectId" -ForegroundColor Green
        } else {
            Write-Error "Could not find function app with managed identity. Please provide -ManagedIdentityObjectId parameter."
            exit 1
        }
    } catch {
        Write-Error "Could not retrieve function app. Please provide -ManagedIdentityObjectId parameter. Error: $_"
        exit 1
    }
}

# Connect to Exchange Online (if not already connected)
$exchangeConnection = Get-ConnectionInformation -ErrorAction SilentlyContinue
if (-not $exchangeConnection) {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
    Connect-ExchangeOnline
}

# Check if group already exists
Write-Host "Checking if group '$GroupName' exists..." -ForegroundColor Yellow
$existingGroup = Get-DistributionGroup -Identity $GroupName -ErrorAction SilentlyContinue

if ($existingGroup) {
    Write-Host "Group '$GroupName' already exists." -ForegroundColor Green
    Write-Host "Group Type: $($existingGroup.GroupType)" -ForegroundColor Gray
    Write-Host "Group mail: $($existingGroup.PrimarySmtpAddress)" -ForegroundColor Gray
} else {
    Write-Host "Creating mail-enabled security group '$GroupName'..." -ForegroundColor Yellow
    
    try {
        # Create as a mail-enabled security group
        New-DistributionGroup `
            -Name $GroupName `
            -Type "Security" `
            -DisplayName $GroupName `
            -Description "Mail-enabled security group for COWG Commanders and Wing Staff to manage shared calendars" `
            -ErrorAction Stop
        
        Write-Host "Group '$GroupName' created successfully." -ForegroundColor Green
        $existingGroup = Get-DistributionGroup -Identity $GroupName
    } catch {
        Write-Error "Failed to create group. Error: $_"
        exit 1
    }
}

# Grant managed identity permission to manage group membership
Write-Host "Configuring managed identity permissions for group management..." -ForegroundColor Yellow

if (-not $ManagedIdentityObjectId) {
    Write-Error "Cannot configure permissions without managed identity object ID."
    exit 1
}

# Get the managed identity service principal
try {
    $managedIdentity = Get-AzADServicePrincipal -ObjectId $ManagedIdentityObjectId -ErrorAction Stop
    Write-Host "Found managed identity service principal: $($managedIdentity.DisplayName)" -ForegroundColor Green
    $managedIdentityEmail = $managedIdentity.ServicePrincipalNames | Select-Object -First 1
} catch {
    Write-Error "Could not find managed identity service principal. Error: $_"
    exit 1
}

try {
    # First, try to get the service principal object
    $managedIdentity = Get-AzADServicePrincipal -ObjectId $ManagedIdentityObjectId -ErrorAction Stop
    
    # Try different identifiers to set as manager
    $identifiersToTry = @(
        $ManagedIdentityObjectId,  # Try object ID directly
        $managedIdentity.AppId,     # Try app ID
        $managedIdentity.DisplayName  # Try display name
    )
    
    $success = $false
    foreach ($identifier in $identifiersToTry) {
        if ($identifier) {
            try {
                Set-DistributionGroup -Identity $GroupName `
                    -ManagedBy $identifier `
                    -ErrorAction Stop
                
                Write-Host "Set managed identity as manager of group '$GroupName' (using: $identifier)." -ForegroundColor Green
                $success = $true
                break
            } catch {
                # Continue to next identifier
                continue
            }
        }
    }
    
    if (-not $success) {
        throw "Could not set any identifier as manager"
    }
    
} catch {
    Write-Error "Failed to set managed identity as manager. Error: $_"
    Write-Host "NOTE: You may need to manually set permissions. Please run with admin:" -ForegroundColor Yellow
    Write-Host "  Set-DistributionGroup -Identity 'COWG-calendar-editors' -ManagedBy '<object-id-or-email>'" -ForegroundColor Gray
}

Write-Host "`nVerifying group configuration..." -ForegroundColor Cyan
$finalGroup = Get-DistributionGroup -Identity $GroupName -ErrorAction Stop
Write-Host "Group Name: $($finalGroup.DisplayName)" -ForegroundColor Gray
Write-Host "Email: $($finalGroup.PrimarySmtpAddress)" -ForegroundColor Gray
Write-Host "Type: $($finalGroup.GroupType)" -ForegroundColor Gray
Write-Host "Mail-Enabled: $($finalGroup.MailEnabled)" -ForegroundColor Gray

Write-Host "`nGroup setup complete!" -ForegroundColor Green

# Check permissions
Write-Host "`nPermissions Configuration:" -ForegroundColor Cyan
try {
    $groupDetails = Get-DistributionGroup -Identity $GroupName -ErrorAction SilentlyContinue
    
    if ($groupDetails.ManagedBy -and $groupDetails.ManagedBy -like "*$ManagedIdentityObjectId*") {
        Write-Host "✓ Managed identity is set as manager of the group." -ForegroundColor Green
    } else {
        Write-Host "⚠ Managed identity is not listed as manager. Current managers:" -ForegroundColor Yellow
        if ($groupDetails.ManagedBy) {
            $groupDetails.ManagedBy | ForEach-Object { Write-Host "  - $_" -ForegroundColor Gray }
        } else {
            Write-Host "  (No managers currently set)" -ForegroundColor Gray
        }
        Write-Host "  If member management fails, ensure the managed identity is set as manager." -ForegroundColor Yellow
    }
} catch {
    Write-Host "⚠ Could not verify manager configuration. Error: $_" -ForegroundColor Yellow
}

Write-Host "`nThe managed identity should now be able to manage group membership." -ForegroundColor Green
Write-Host "You can now run the DLOpsQuals function with the calendar editors sync enabled.`n" -ForegroundColor Cyan
