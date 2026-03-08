# Shared functions and module loading for CAPWATCH Azure Functions

# Load required modules at runtime (for deployment size optimization)
. "$PSScriptRoot\Load-Modules.ps1"

# Initialize runtime modules from Azure Storage
. "$PSScriptRoot\Load-Modules.ps1"

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

# Function: Get-DryRunMode
# Purpose: Determines if the function should run in dry-run mode (preview changes without executing)
# Returns: $true if in dry-run mode, $false if in execute mode
function Get-DryRunMode {
    # Check environment variables (case-insensitive)
    $executeValue = $env:EXECUTE
    
    # Default to dry-run mode (safe) if not specified
    if ([string]::IsNullOrEmpty($executeValue)) {
        return $true
    }
    
    # Parse the EXECUTE variable - false/no/dryrun = dry-run mode, true/yes = execute mode
    return -not ($executeValue -match '^(true|yes|1)$')
}

# Function: Write-OperationLog
# Purpose: Logs an operation with DRY-RUN indicator if in dry-run mode
# Usage: Write-OperationLog -Operation "Creating Team" -Details "Team: MyTeam" -IsPreview $true
function Write-OperationLog {
    param (
        [string]$Operation,
        [string]$Details = "",
        [bool]$IsExecuting = $false
    )
    
    $dryRun = Get-DryRunMode
    
    if ($dryRun) {
        $prefix = "🔍 [DRY-RUN]"
    } else {
        $prefix = "⚡ [EXECUTING]"
    }
    
    if ($Details) {
        Write-Log "$prefix $Operation - $Details"
    } else {
        Write-Log "$prefix $Operation"
    }
}

# Function: Test-ExecutionMode
# Purpose: Tests whether an operation should actually execute
# Usage: if (Test-ExecutionMode) { /* make changes */ }
function Test-ExecutionMode {
    $dryRun = Get-DryRunMode
    return -not $dryRun
}

# Function: GetAllUsers
# Purpose: Retrieves all users from Microsoft Graph API.
function GetAllUsers {
    param (
        [string]$SelectFields = "mail,displayName,officeLocation,companyName,employeeId,id,employeeType,jobTitle,department"
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
function GetDeletedUsers {
    # Define the API endpoint to query deleted users
    $deletedUsersUri = "https://graph.microsoft.com/beta/directory/deletedItems/microsoft.graph.user"
    # Retrieve deleted users
    $deletedUsers = @()
    do {
        $apiResponse = Invoke-MgGraphRequest -Method GET -Uri $deletedUsersUri
        $deletedUsers += $apiResponse.value
        $deletedUsersUri = $apiResponse.'@odata.nextLink'
    } while ($deletedUsersUri)
    return $deletedUsers
}

# Function: Remove-OldLogFiles
# Purpose: Removes log files older than specified number of days from a directory
function Remove-OldLogFiles {
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

# Function: Remove-ExpiredMemberAccounts
# Purpose: Removes expired member accounts and related Exchange contacts
function Remove-ExpiredMemberAccounts {
    Write-Log "Starting expired member account deletion process..."
    
    # Set working directory to folder with all CAPWATCH CSV Text Files
    $CAPWATCHDATADIR = "$($env:HOME)\data\CAPWatch"
    
    # Import the CSV files
    $expiredMembers = Import-Csv "$($CAPWATCHDATADIR)\Member.txt" -ErrorAction Stop | Where-Object { $_.MbrStatus -eq "EXPIRED" }
    $contacts = Import-Csv "$($CAPWATCHDATADIR)\MbrContact.txt" -ErrorAction Stop
    
    # Get all current users from Microsoft Graph
    $allUsers = Get-MgUser -All
    
    # Array to track deleted members for notifications
    $deletedMembersList = @()
    
    # Process each expired member
    foreach ($expiredMember in $expiredMembers) {
        $capid = $expiredMember.CAPID
        $parentCAPID = $capid + "P"
        
        # Find the member's account in Azure AD
        $memberAccount = $allUsers | Where-Object { $_.officeLocation -eq $capid }
        $parentAccount = $allUsers | Where-Object { $_.officeLocation -eq $parentCAPID }
        
        # Get the member's email from contacts
        $memberEmail = ($contacts | Where-Object { $_.CAPID -eq $capid -and $_.Contact -match '^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$' -and $_.Priority -eq "PRIMARY" } | Select-Object -First 1).Contact
        
        # Helper: check if a user exists by id (returns $true/$false). Treat 404 as not exists.
        function Test-UserExistsById {
            param([string]$id)
            try {
                $checkUri = "https://graph.microsoft.com/v1.0/users/$id"
                Invoke-MgGraphRequest -Method GET -Uri $checkUri -ErrorAction Stop | Out-Null
                return $true
            } catch {
                $err = $_.ToString()
                if ($err -match 'Request_ResourceNotFound' -or $err -match '404 Not Found') {
                    return $false
                }
                # For other errors, rethrow to be handled by caller
                throw
            }
        }

        # Delete the member's account
        if ($memberAccount) {
            try {
                $userId = $memberAccount.id
                if (-not (Test-UserExistsById -id $userId)) {
                    Write-Log "Member account $userId not found before delete (already removed). Skipping."
                } else {
                    $uri = "https://graph.microsoft.com/v1.0/users/$userId"
                    Invoke-MgGraphRequest -Method DELETE -Uri $uri
                    Write-Log "Deleted member account: $($memberAccount.displayName) ($($memberAccount.mail)) with CAPID: $capid."
                }
                
                # Add to deleted members list for notification
                $deletedMembersList += [PSCustomObject]@{
                    NameFirst = $expiredMember.NameFirst
                    NameLast = $expiredMember.NameLast
                    Grade = $expiredMember.Rank
                    CAPID = $capid
                    Email = if ($memberAccount.mail) { $memberAccount.mail } else { $memberEmail }
                    Unit = $expiredMember.companyName
                    Type = "Member"
                }
                
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
                $errStr = $_.ToString()
                Write-Log "Failed to delete member account $($memberAccount.id): $errStr"
            }
        }
        
        # Delete the parent's guest account
        if ($parentAccount) {
            try {
                $parentId = $parentAccount.id
                if (-not (Test-UserExistsById -id $parentId)) {
                    Write-Log "Parent account $parentId not found before delete (already removed). Skipping."
                } else {
                    $uri = "https://graph.microsoft.com/v1.0/users/$parentId"
                    Invoke-MgGraphRequest -Method DELETE -Uri $uri
                    Write-Log "Deleted parent account: $($parentAccount.displayName) ($($parentAccount.mail)) with CAPID: $parentCAPID."
                }
                
                # Get parent email from contacts
                $parentEmail = ($contacts | Where-Object { $_.CAPID -eq $capid -and $_.Type -eq "CADET PARENT EMAIL" } | Select-Object -First 1).Contact
                
                # Add to deleted members list for notification
                $deletedMembersList += [PSCustomObject]@{
                    NameFirst = $expiredMember.NameFirst
                    NameLast = "$($expiredMember.NameLast) (Parent)"
                    Grade = "$($expiredMember.Rank) PARENT"
                    CAPID = $parentCAPID
                    Email = if ($parentAccount.mail) { $parentAccount.mail } else { $parentEmail }
                    Unit = $expiredMember.companyName
                    Type = "Parent"
                }
                
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
                $errStr = $_.ToString()
                Write-Log "Failed to delete parent account $($parentAccount.id): $errStr"
            }
        }
    }
    
    # List of CAPIDs to exclude from deletion (exceptions)
    $excludeCAPIDs = @('360390', '185483', '672934') # Add CAPIDs to exclude here
    
    # Log all O365 members whose CAPIDs don't exist in $members and CAPID is not 999999 (with 'P' suffix handled)
    $memberCAPIDs = Import-Csv "$($CAPWATCHDATADIR)\Member.txt" -ErrorAction Stop | Where-Object { $_.MbrStatus -eq "ACTIVE" } | ForEach-Object { $_.CAPID }
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
            Write-Log "DisplayName: $($user.displayName), Email: $($user.mail), CAPID: $($user.officeLocation), Unit: $($user.companyName)"
                try {
                    $userId = $user.id
                    if (-not (Test-UserExistsById -id $userId)) {
                        Write-Log "O365 account $userId not found before delete (already removed). Skipping."
                    } else {
                        $uri = "https://graph.microsoft.com/beta/users/$userId"
                        Invoke-MgGraphRequest -Method DELETE -Uri $uri
                        Write-Log "Deleted O365 account: $($user.displayName) ($($user.mail)), CAPID: $($user.officeLocation), Unit: $($user.companyName)"
                    }
                
                # Add to deleted members list for notification
                # Extract names from displayName (format: "First Last, Rank")
                $nameParts = $user.displayName -split ', '
                $fullName = $nameParts[0]
                $rank = if ($nameParts.Count -gt 1) { $nameParts[1] } else { "Unknown" }
                $nameComponents = $fullName -split ' '
                $firstName = $nameComponents[0]
                $lastName = if ($nameComponents.Count -gt 1) { $nameComponents[1..($nameComponents.Count-1)] -join ' ' } else { "" }
                
                # Determine unit from companyName only
                $unit = "Unknown"
                if ($user.companyName -and $user.companyName -match 'CO-(.+)') {
                    $unit = $matches[1]  # Keep exact format from companyName
                }
                Write-Log "User: $($user.displayName), CompanyName: $($user.companyName), Extracted Unit: $unit"
                
                $deletedMembersList += [PSCustomObject]@{
                    NameFirst = $firstName
                    NameLast = $lastName
                    Grade = $rank
                    CAPID = $user.officeLocation
                    Email = if ($user.mail) { $user.mail } else { $user.userPrincipalName }
                    Unit = $unit
                    Type = "Inactive"
                }
            } catch {
                $errStr = $_.ToString()
                if ($errStr -match 'Request_ResourceNotFound' -or $errStr -match '404 Not Found') {
                    Write-Log "O365 account $($user.id) not found when attempting delete (already removed)."
                } else {
                    Write-Log "Failed to delete O365 account: $($user.displayName) ($($user.mail)), CAPID: $($user.officeLocation). Error: $_"
                }
            }
        }
    } else {
        Write-Log "All O365 users have CAPIDs present in the CAPWATCH members list or are 999999 (with 'P' suffix handled)."
    }
    
    # Send notification emails about deleted expired members
    Write-Log "Sending notification emails about deleted expired members..."
    Send-ExpiredMembersNotification -deletedMembers $deletedMembersList -allUsers $allUsers
    
    Write-Log "Expired member account deletion process completed."
}

# Function: Get-UnitNotificationEmails
# Purpose: Get notification emails for a unit's commanders and recruiting officer
function Get-UnitNotificationEmails {
    param (
        [string]$unit,
        [array]$allUsers
    )
    $emails = @()
    foreach ($user in $allUsers) {
        if ($user.companyName -match $unit -and $user.department -match '(PA|EX)') {
            if ($user.mail) { $emails += $user.mail }
        }
    }
    $emails = $emails | Select-Object -Unique
    return $emails
}

# Function: Send-ExpiredMembersNotification
# Purpose: Send email notifications about expired members that were removed
function Send-ExpiredMembersNotification {
    param (
        [array]$deletedMembers,
        [array]$allUsers
    )
    
    if ($deletedMembers.Count -eq 0) {
        Write-Log "No expired members to report."
        return
    }
    
    Write-Log "Found $($deletedMembers.Count) deleted members to report."
    
    # Group deleted members by unit
    $deletedByUnit = $deletedMembers | Group-Object -Property Unit
    
    foreach ($unitGroup in $deletedByUnit) {
        $unit = $unitGroup.Name
        $unitMembers = $unitGroup.Group
        
        Write-Log "Processing notifications for unit: $unit with $($unitMembers.Count) deleted members"
        
        # Build the toRecipients array
        $toRecipients = @()
        # Always add mike.schulte@cowg.cap.gov
        $toRecipients += "mike.schulte@cowg.cap.gov"
        
        # Add the unit's recruiting distribution list
        # Unit format is like "CO-173", so recruiting DL is "co-173-recruiting@cowg.cap.gov"
        $recruitingDL = "$(($unit).ToLower())-recruiting@cowg.cap.gov"
        if ($recruitingDL -and $recruitingDL -ne "mike.schulte@cowg.cap.gov") {
            $toRecipients += $recruitingDL
        }

        Write-Log "Sending deletion notification to: $($toRecipients -join ', ')"
        
        # Build the member list HTML table
        $memberTableRows = ""
        foreach ($member in $unitMembers) {
            $memberTableRows += @"
      <tr>
        <td style='padding: 8px; border: 1px solid #ddd;'>$($member.NameFirst) $($member.NameLast)</td>
        <td style='padding: 8px; border: 1px solid #ddd;'>$($member.Grade)</td>
        <td style='padding: 8px; border: 1px solid #ddd;'>$($member.CAPID)</td>
        <td style='padding: 8px; border: 1px solid #ddd;'>$($member.Email)</td>
        <td style='padding: 8px; border: 1px solid #ddd;'>$($member.Type)</td>
      </tr>
"@
        }
        
        # Convert email strings to proper Graph API format
        $graphToRecipients = @()
        foreach ($email in $toRecipients) {
            $graphToRecipients += @{
                emailAddress = @{
                    address = $email
                }
            }
        }
        
        try {
            $userPrincipalName = "cowg_it_helpdesk@cowg.cap.gov" # Use a service account or shared mailbox with Mail.Send permission
            
            $mailBody = @{
                message = @{
                    subject = "Expired Members Removed from CO-$unit"
                    body = @{
                        contentType = "HTML"
                        content = @"
<html>
  <body style='font-family: Arial, sans-serif; color: #222;'>
    <div style='text-align: center; margin-bottom: 20px;'>
      <img src='https://cowg.cap.gov/media/websites/COWG_T_7665FADF8B38C.PNG' alt='COWG Logo' style='max-width: 200px;'/>
    </div>
    <h2 style='color: #003366;'>Expired Members Removed from CO-$unit</h2>
    <p>The following members have been removed from CO-$unit because their membership has expired in CAPWATCH or they are no longer active:</p>
    <table style='margin: 20px auto; border-collapse: collapse; width: 90%;'>
      <thead>
        <tr style='background-color: #f2f2f2;'>
          <th style='padding: 8px; border: 1px solid #ddd; text-align: left;'>Name</th>
          <th style='padding: 8px; border: 1px solid #ddd; text-align: left;'>Grade</th>
          <th style='padding: 8px; border: 1px solid #ddd; text-align: left;'>CAPID</th>
          <th style='padding: 8px; border: 1px solid #ddd; text-align: left;'>Email</th>
          <th style='padding: 8px; border: 1px solid #ddd; text-align: left;'>Type</th>
        </tr>
      </thead>
      <tbody>
$memberTableRows
      </tbody>
    </table>
    <p><strong>Total removed:</strong> $($unitMembers.Count) member(s)</p>
    <p style='font-size: 0.9em; color: #888; margin-top: 30px;'>This is an automated notification from the COWG IT Team. These accounts have been permanently deleted from Azure AD and Exchange Online.</p>
  </body>
</html>
"@
                    }
                    toRecipients = $graphToRecipients
                }
                saveToSentItems = $false
            } | ConvertTo-Json -Depth 4
            
            $uri = "https://graph.microsoft.com/v1.0/users/$userPrincipalName/sendMail"
            Invoke-MgGraphRequest -Method POST -Uri $uri -Body $mailBody -ContentType "application/json"
            Write-Log "Expired members notification email sent for unit CO-$unit to: $($toRecipients -join ', ')"
        } catch {
            Write-Log "Failed to send expired members notification email for unit CO-${unit}: $_"
            Write-Log "Error details: $($_.Exception.Message)"
        }
    }
}

# Function: Get-CosmosDbConnection
# Purpose: Retrieves Cosmos DB connection configuration from environment variables.
function Get-CosmosDbConnection {
    param (
        [string]$ConnectionString = $env:CosmosDbConnectionString,
        [string]$Database = $env:CosmosDbDatabase,
        [string]$Container = $env:CosmosDbContainer
    )
    
    if (-not $ConnectionString -or -not $Database -or -not $Container) {
        Write-Log "Warning: Cosmos DB configuration incomplete. ConnectionString: $(if($ConnectionString) {'Set'} else {'Not Set'}), Database: $Database, Container: $Container"
    }
    
    return @{
        ConnectionString = $ConnectionString
        Database = $Database
        Container = $Container
        Endpoint = "https://cosmos-oflight-coordinator.documents.azure.com"
    }
}

# Function: Save-CosmosDbItem
# Purpose: Saves an item (document) to Azure Cosmos DB using HTTP REST API with proper auth and retry logic.
function Save-CosmosDbItem {
    param (
        [Parameter(Mandatory=$true)]
        [object]$Item,
        [string]$ConnectionString = $env:CosmosDbConnectionString,
        [string]$Database = $env:CosmosDbDatabase,
        [string]$Container = $env:CosmosDbContainer,
        [int]$MaxRetries = 3,
        [int]$InitialDelayMs = 500
    )
    
    # Transient error patterns that should be retried
    $transientErrorPatterns = @(
        'connection.*timeout',
        'connection.*failed',
        'did not properly respond',
        'host has failed to respond',
        'request.*timeout',
        'connection reset',
        'The remote name could not be resolved'
    )
    
    # Helper function to check if error is transient
    function Test-IsTransientError {
        param([string]$ErrorMessage)
        foreach ($pattern in $transientErrorPatterns) {
            if ($ErrorMessage -imatch $pattern) {
                return $true
            }
        }
        return $false
    }
    
    $attempt = 0
    
    while ($attempt -le $MaxRetries) {
        try {
            $attempt++
            
            # Ensure item has required id field
            if (-not $Item.id) {
                $Item | Add-Member -MemberType NoteProperty -Name "id" -Value ([guid]::NewGuid().ToString()) -ErrorAction SilentlyContinue
            }
            
            # Add timestamp if not present
            if (-not $Item.timestamp) {
                $Item | Add-Member -MemberType NoteProperty -Name "timestamp" -Value (Get-Date -Format o) -ErrorAction SilentlyContinue
            }
            
            # Parse connection string
            $connStringParts = @{}
            $ConnectionString -split ';' | Where-Object { $_ -match '=' } | ForEach-Object {
                $key, $value = $_ -split '=', 2
                $connStringParts[$key.Trim()] = $value.Trim()
            }
            
            $endpoint = $connStringParts['AccountEndpoint']
            $key = $connStringParts['AccountKey']
            
            if (-not $endpoint -or -not $key) {
                Write-Log "Failed to parse Cosmos DB connection string"
                return $false
            }
            
            # Remove trailing slash if present
            $endpoint = $endpoint.TrimEnd('/')
            
            # Build URI for upsert
            $uri = "$endpoint/dbs/$Database/colls/$Container/docs"
            
            # Generate auth header - Cosmos DB REST API requires specific format
            $verb = "post"
            $resourceType = "docs"
            $resourceId = "dbs/$Database/colls/$Container"
            $date = [DateTime]::UtcNow.ToString('r')
            
            # Build the string to sign (lowercase verb, resourceType, resourceId, and date per Cosmos DB spec)
            $stringToSign = "$verb`n$resourceType`n$resourceId`n$($date.ToLowerInvariant())`n`n"
            
            # Create HMAC SHA256 signature
            $hmacsha = New-Object System.Security.Cryptography.HMACSHA256
            $hmacsha.Key = [System.Convert]::FromBase64String($key)
            $hashBytes = $hmacsha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($stringToSign))
            $signature = [System.Convert]::ToBase64String($hashBytes)
            
            # Build the authorization token and URL-encode the ENTIRE string (per Cosmos DB docs)
            $authString = "type=master&ver=1.0&sig=$signature"
            $authToken = [System.Web.HttpUtility]::UrlEncode($authString)
            
            # Get partition key value from the item (CAPID is our partition key)
            $partitionKeyValue = $Item.CAPID
            
            $headers = @{
                "Authorization"                  = $authToken
                "x-ms-date"                      = $date
                "x-ms-version"                   = "2020-07-15"
                "x-ms-documentdb-is-upsert"      = "true"
                "x-ms-documentdb-partitionkey"   = "[`"$partitionKeyValue`"]"
            }
            
            $body = $Item | ConvertTo-Json -Depth 10
            
            Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $body -ContentType "application/json" -ErrorAction Stop | Out-Null
            
            Write-Log "Successfully saved item to Cosmos DB: $($Item.id)"
            return $true
            
        } catch {
            $errorMessage = $_.Exception.Message
            $isTransient = Test-IsTransientError -ErrorMessage $errorMessage
            
            if ($attempt -le $MaxRetries -and $isTransient) {
                # Calculate exponential backoff delay: 500ms, 1s, 2s, etc.
                $delayMs = $InitialDelayMs * [Math]::Pow(2, $attempt - 1)
                Write-Log "Transient error on attempt $attempt/$MaxRetries for item $($Item.id). Retrying in ${delayMs}ms. Error: $errorMessage"
                Start-Sleep -Milliseconds $delayMs
            } else {
                # Non-transient error or max retries exceeded
                if ($attempt -gt $MaxRetries) {
                    Write-Log "Failed to save item to Cosmos DB after $MaxRetries retries: $($Item.id). Error: $errorMessage"
                } else {
                    Write-Log "Failed to save item to Cosmos DB (non-transient error): $($Item.id). Error: $errorMessage"
                }
                return $false
            }
        }
    }
    
    # Should not reach here, but just in case
    Write-Log "Failed to save item to Cosmos DB: $($Item.id) after all retry attempts"
    return $false
}

# Function: Invoke-CosmosDbQuery
# Purpose: Executes a query against Azure Cosmos DB.
function Invoke-CosmosDbQuery {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Query,
        [string]$ConnectionString = $env:CosmosDbConnectionString,
        [string]$Database = $env:CosmosDbDatabase,
        [string]$Container = $env:CosmosDbContainer
    )
    
    try {
        Write-Log "Query function called but not implemented - returning empty results"
        return @()
    } catch {
        Write-Log "Failed to query Cosmos DB. Error: $($_.Exception.Message)"
        return @()
    }
}

# Function: Get-CosmosDbItem
# Purpose: Retrieves a specific item from Cosmos DB by ID.
function Get-CosmosDbItem {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ItemId,
        [string]$ConnectionString = $env:CosmosDbConnectionString,
        [string]$Database = $env:CosmosDbDatabase,
        [string]$Container = $env:CosmosDbContainer
    )
    
    try {
        # For now, always return null to trigger upsert
        # In production, this could query Cosmos DB
        return $null
    } catch {
        Write-Log "Failed to retrieve item from Cosmos DB. Error: $($_.Exception.Message)"
        return $null
    }
}