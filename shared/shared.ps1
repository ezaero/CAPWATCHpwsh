# Function: Write-Log
# Purpose: Logs messages to a file and outputs them to the console.
function Write-Log {
    param (
        [string]$Message
    )
    $LogFile = "$env:HOME\logs\script_log_$(Get-Date -Format 'yyyy-MM-dd').txt"
    # Ensure the directory exists
    $logDirectory = [System.IO.Path]::GetDirectoryName($LogFile)
    if (-not (Test-Path -Path $logDirectory)) {
        New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
    }

    # Write the log message
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogFile -Value "$timestamp - $Message"
    Write-Host "$timestamp - $Message"
}

# Function: GetAllUsers
# Purpose: Retrieves all users from Microsoft Graph API.
function GetAllUsers {
    param (
        [string]$SelectFields = "mail,displayName,officeLocation,companyName,employeeId,id,employeeType,jobTitle"
    )

    $allUsers = @()
    $uri = "https://graph.microsoft.com/beta/users?$select=$SelectFields"
    do {
        try {
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri
            $allUsers += $response.value
            $uri = $response.'@odata.nextLink'
        } catch {
            Write-Log "Failed to fetch users from Microsoft Graph API. Error: $($_.Exception.Message)"
            break
        }
    } while ($uri)
    return $allUsers
}