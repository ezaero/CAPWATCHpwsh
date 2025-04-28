# Input bindings are passed in via param block.
param($Timer)

# Connect to Microsoft Graph
$MSGraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token

Connect-MgGraph -AccessToken $MSGraphAccessToken -NoWelcome
# Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -ManagedIdentity -Organization COCivilAirPatrol.onmicrosoft.com

$logFile = 

function Write-Log {
    param (
        [string]$Message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logFile -Value "$timestamp - $Message"
    Write-Host "$timestamp - $Message"
}

function Delete-OldLogFiles {
    param (
        [string]$DirectoryPath, # The directory containing the log files
        [int]$DaysOld = 30      # The age threshold in days (default is 30 days)
    )

    # Check if the directory exists
    if (-not (Test-Path -Path $DirectoryPath)) {
        Write-Host "Directory '$DirectoryPath' does not exist. Exiting." -ForegroundColor Red
        return
    }

    # Get the current date
    $currentDate = Get-Date

    # Find all log files (*.txt) older than the specified number of days
    $oldFiles = Get-ChildItem -Path $DirectoryPath -Filter "*.txt" -File | Where-Object {
        ($currentDate - $_.LastWriteTime).Days -gt $DaysOld
    }

    # Delete the old files
    foreach ($file in $oldFiles) {
        try {
            Remove-Item -Path $file.FullName -Force
            Write-Host "Deleted log file: $($file.FullName)" -ForegroundColor Green
        } catch {
            Write-Host "Failed to delete log file: $($file.FullName). Error: $_" -ForegroundColor Red
        }
    }

    # Log the result
    if ($oldFiles.Count -eq 0) {
        Write-Host "No log files older than $DaysOld days were found in '$DirectoryPath'." -ForegroundColor Yellow
    } else {
        Write-Host "Deleted $($oldFiles.Count) log file(s) older than $DaysOld days from '$DirectoryPath'." -ForegroundColor Green
    }
}

Delete-OldLogFiles
