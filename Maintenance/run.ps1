# Input bindings are passed in via param block.
param($Timer)

# Set working directory to folder with all CAPWATCH CSV Text Files
$CAPWATCHDATADIR = "$($env:HOME)\data\CAPWatch"
Push-Location $CAPWATCHDATADIR

# Include shared Functions
. "$PSScriptRoot\..\shared\shared.ps1"

# Connect to Microsoft Graph
$MSGraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token

Connect-MgGraph -AccessToken $MSGraphAccessToken -NoWelcome
Connect-ExchangeOnline -ManagedIdentity -Organization $env:EXCHANGE_ORGANIZATION

function Delete-OldLogFiles {
    param (
        [string]$DirectoryPath, # The directory containing the log files
        [int]$DaysOld = 30      # The age threshold in days (default is 30 days)
    )

    # Check if the directory exists
    if (-not (Test-Path -Path $DirectoryPath)) {
        Write-Log "Directory '$DirectoryPath' does not exist. Exiting."
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
            Write-Log "Deleted log file: $($file.FullName)"
        } catch {
            Write-Log "Failed to delete log file: $($file.FullName). Error: $_"
        }
    }

    # Log the result
    if ($oldFiles.Count -eq 0) {
        Write-Log "No log files older than $DaysOld days were found in '$DirectoryPath'."
    } else {
        Write-Log "Deleted $($oldFiles.Count) log file(s) older than $DaysOld days from '$DirectoryPath'."
    }
}

function Delete-ExpiredMemberAccounts {
    Write-Log "Starting expired member account deletion process..."
    
    # Import the CSV files
    $members = Import-Csv "$($CAPWATCHDATADIR)\Member.txt" -ErrorAction Stop | Where-Object { $_.MbrStatus -eq "ACTIVE" }
    $expiredMembers = Import-Csv "$($CAPWATCHDATADIR)\Member.txt" -ErrorAction Stop | Where-Object { $_.MbrStatus -eq "EXPIRED" }
    
    # Get all users from Azure AD
    $allUsers = @()
    $uri = "https://graph.microsoft.com/beta/users?`$select=mail,displayName,officeLocation,companyName,employeeId"
    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
        $allUsers += $response.value
        $uri = $response.'@odata.nextLink'
    } while ($uri)
    
    Write-Log "Processing $($expiredMembers.Count) expired members for account deletion..."

    # Loop through each expired member
    foreach ($expiredMember in $expiredMembers) {
        $capid = $expiredMember.CAPID
        $parentCAPID = "$capid`P" # Parent's CAPID is CAPID + "P"

        # Find the member's account in Azure AD
        $memberAccount = $allUsers | Where-Object { $_.officeLocation -eq $capid }
        $parentAccount = $allUsers | Where-Object { $_.officeLocation -eq $parentCAPID }

        # Delete the member's account
        if ($memberAccount) {
            try {
                $uri = "https://graph.microsoft.com/v1.0/users/$($memberAccount.id)"
                Invoke-MgGraphRequest -Method DELETE -Uri $uri
                Write-Log "Deleted member account: $($memberAccount.displayName) ($($memberAccount.mail)) with CAPID: $capid."
                # Also check if they are an Exchange contact and delete that
                $contactEmail = $memberAccount.mail
                if (-not $contactEmail) { $contactEmail = $memberAccount.userPrincipalName }
                if ($contactEmail -and $contactEmail -match '^[\w\.-]+@([\w-]+\.)+[\w-]{2,}$') {
                    $existingContact = Get-MailContact -Filter "ExternalEmailAddress -eq '$contactEmail'" -ErrorAction SilentlyContinue
                    if ($existingContact) {
                        try {
                            Remove-MailContact -Identity $existingContact.Identity -Confirm:$false
                            Write-Log "Deleted Exchange mail contact for $contactEmail."
                        } catch {
                            Write-Log ("Failed to delete Exchange mail contact for {0}: {1}" -f $contactEmail, $_)
                        }
                    }
                }
            } catch {
                Write-Log "Failed to delete member account: $($memberAccount.displayName) ($($memberAccount.mail)). Error: $_"
            }
        }

        # Delete the parent's guest account
        if ($parentAccount) {
            try {
                $uri = "https://graph.microsoft.com/v1.0/users/$($parentAccount.id)"
                Invoke-MgGraphRequest -Method DELETE -Uri $uri
                Write-Log "Deleted parent account: $($parentAccount.displayName) ($($parentAccount.mail)) with CAPID: $parentCAPID."
                # Also check if they are an Exchange contact and delete that
                $contactEmail = $parentAccount.mail
                if (-not $contactEmail) { $contactEmail = $parentAccount.userPrincipalName }
                if ($contactEmail -and $contactEmail -match '^[\w\.-]+@([\w-]+\.)+[\w-]{2,}$') {
                    $existingContact = Get-MailContact -Filter "ExternalEmailAddress -eq '$contactEmail'" -ErrorAction SilentlyContinue
                    if ($existingContact) {
                        try {
                            Remove-MailContact -Identity $existingContact.Identity -Confirm:$false
                            Write-Log "Deleted Exchange mail contact for $contactEmail."
                        } catch {
                            Write-Log ("Failed to delete Exchange mail contact for {0}: {1}" -f $contactEmail, $_)
                        }
                    }
                }
            } catch {
                Write-Log "Failed to delete parent account: $($parentAccount.displayName) ($($parentAccount.mail)). Error: $_"
            }
        }
    }
    
    # List of CAPIDs to exclude from deletion (exceptions)
    $excludeCAPIDs = @('360390', '185483', '672934') # Add CAPIDs to exclude here

    # Log all O365 members whose CAPIDs don't exist in $members and CAPID is not 999999 (with 'P' suffix handled)
    $memberCAPIDs = $members | ForEach-Object { $_.CAPID }
    $missingCAPIDUsers = @()
    foreach ($user in $allUsers) {
        if ($user.officeLocation) {
            $capidToCheck = $user.officeLocation
            if ($capidToCheck -match '^\d+P$') {
                $capidToCheck = $capidToCheck.Substring(0, $capidToCheck.Length - 1)
            }
            if (($memberCAPIDs -notcontains $capidToCheck) -and $capidToCheck -ne '999999' -and ($excludeCAPIDs -notcontains $capidToCheck)) {
                $missingCAPIDUsers += $user
            }
        }
    }
    if ($missingCAPIDUsers.Count -gt 0) {
        Write-Log "O365 users whose CAPIDs do not exist in CAPWATCH members list and are not 999999 (with 'P' suffix handled):"
        foreach ($user in $missingCAPIDUsers) {
            Write-Log "DisplayName: $($user.displayName), Email: $($user.mail), CAPID: $($user.officeLocation)"
            try {
                $uri = "https://graph.microsoft.com/v1.0/users/$($user.id)"
                Invoke-MgGraphRequest -Method DELETE -Uri $uri
                Write-Log "Deleted O365 account: $($user.displayName) ($($user.mail)), CAPID: $($user.officeLocation)"
            } catch {
                Write-Log "Failed to delete O365 account: $($user.displayName) ($($user.mail)), CAPID: $($user.officeLocation). Error: $_"
            }
        }
    } else {
        Write-Log "All O365 users have CAPIDs present in the CAPWATCH members list or are 999999 (with 'P' suffix handled)."
    }
    
    Write-Log "Expired member account deletion process completed."
}

# Always delete old log files
Delete-OldLogFiles -DirectoryPath "$env:HOME\logs"

# Run account deletion maintenance (now scheduled by function.json)
Write-Log "Running monthly account deletion maintenance..."
Delete-ExpiredMemberAccounts
