# Test script for OFlightMetrics function - local testing without timer trigger
# This bypasses the timer trigger and tests the core logic directly

Write-Host "==================== OFlightMetrics Test Script ====================" -ForegroundColor Cyan
Write-Host ""

# Create logs directory if it doesn't exist
$logsDir = "$env:HOME\logs"
if (-not (Test-Path -Path $logsDir)) {
    New-Item -ItemType Directory -Path $logsDir -Force | Out-Null
    Write-Host "Created logs directory: $logsDir" -ForegroundColor Gray
}

# Load environment variables from local.settings.json
Write-Host "Loading environment variables from local.settings.json..." -ForegroundColor Yellow
$localSettingsPath = "$PSScriptRoot\local.settings.json"
if (Test-Path $localSettingsPath) {
    $localSettings = Get-Content $localSettingsPath | ConvertFrom-Json
    foreach ($key in $localSettings.Values.PSObject.Properties.Name) {
        $value = $localSettings.Values.$key
        if ($value -and -not $key.StartsWith('_comment')) {
            [System.Environment]::SetEnvironmentVariable($key, $value, 'Process')
        }
    }
    Write-Host "✅ Environment variables loaded" -ForegroundColor Green
} else {
    Write-Host "⚠️  Warning: local.settings.json not found" -ForegroundColor Yellow
}
Write-Host ""

# Include shared Functions
. "$PSScriptRoot\shared\shared.ps1"

# Test Cosmos DB connection
Write-Host "Testing Cosmos DB configuration..." -ForegroundColor Yellow
$cosmosConfig = Get-CosmosDbConnection
if (-not $cosmosConfig.ConnectionString -or -not $cosmosConfig.Database) {
    Write-Host "❌ Error: Cosmos DB configuration incomplete" -ForegroundColor Red
    Write-Host "   Please ensure CosmosDbConnectionString and CosmosDbDatabase environment variables are set" -ForegroundColor Red
    exit 1
}
Write-Host "✅ Cosmos DB configuration found" -ForegroundColor Green
Write-Host "   Database: $($cosmosConfig.Database)" -ForegroundColor Gray
Write-Host ""

# Test Azure connection and authenticate if needed
Write-Host "Testing Azure connection..." -ForegroundColor Yellow
try {
    # Try to get MSGraph token first
    $tokenTest = $null
    try {
        $tokenTest = Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
    } catch {
        # Token request failed
    }

    if (-not $tokenTest) {
        Write-Host "⚠️  Need to authenticate with Microsoft Graph scope..." -ForegroundColor Yellow
        Write-Host "   Opening browser for authentication..." -ForegroundColor Gray

        # Disconnect any existing session first
        Disconnect-AzAccount -ErrorAction SilentlyContinue | Out-Null

        # Connect with proper scope
        Connect-AzAccount -AuthScope MicrosoftGraphEndpointResourceId
        Write-Host "✅ Connected to Azure with Microsoft Graph scope" -ForegroundColor Green
    } else {
        $context = Get-AzContext
        Write-Host "✅ Already connected to Azure as $($context.Account.Id)" -ForegroundColor Green
    }
} catch {
    Write-Host "❌ Failed to connect to Azure" -ForegroundColor Red
    Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
Write-Host ""

# Test Microsoft Graph connection
Write-Host "Testing Microsoft Graph connection..." -ForegroundColor Yellow
try {
    $MSGraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token
    Connect-MgGraph -AccessToken $MSGraphAccessToken -NoWelcome
    Write-Host "✅ Successfully authenticated to Microsoft Graph" -ForegroundColor Green
} catch {
    Write-Host "❌ Failed to authenticate to Microsoft Graph" -ForegroundColor Red
    Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
Write-Host ""

# Execute the metrics function
Write-Host "Executing OFlightMetrics function..." -ForegroundColor Yellow
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan

# Call the actual function script
& "$PSScriptRoot\OFlightMetrics\run.ps1" -Timer @{}

Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "✅ OFlightMetrics test completed!" -ForegroundColor Green
Write-Host ""
