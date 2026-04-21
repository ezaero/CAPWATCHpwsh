<#
.SYNOPSIS
    Synchronizes CAPWATCH data with Microsoft Entra ID (Azure AD) and ensures accurate user information in O365.

.DESCRIPTION
    This script performs the following tasks:
    1. Connects to Microsoft Graph API using the Microsoft Graph PowerShell SDK.
    2. Imports data from CAPWATCH CSV files (`MbrContact.txt`, `Member.txt`, and `DutyPosition.txt`).
    3. Combines the data from the CSV files into a unified dataset for processing.
    4. Deduplicates member records by email address, preferring non-PARENT accounts when duplicates exist.
    5. Compares the CAPWATCH data with existing Microsoft Entra ID (Azure AD) users to:
        - Identify users to be added as O365 guest accounts.
        - Identify users to be removed from O365.
        - Ensure all O365 accounts have the correct CAPID, duty positions, and unit information.
    6. Creates O365 guest accounts for users missing in Azure AD.
        - Prevents duplicate accounts by checking if an email already exists in Azure AD (safety check).
        - Logs any attempts to create duplicate accounts for informational and auditing purposes.
    7. Updates existing O365 accounts with CAPID, duty positions, and unit information.
    8. Removes users from O365 who are no longer in the CAPWATCH data.
    9. Identifies and logs users with duplicate display names in Azure AD.
    10. Exports users with missing CAPIDs and logs all actions for auditing purposes.

.PARAMETER contactsFile
    Path to the `MbrContact.txt` file containing contact information.

.PARAMETER memberFile
    Path to the `Member.txt` file containing member information.

.PARAMETER dutyPositionFile
    Path to the `DutyPosition.txt` file containing duty position information.

.PARAMETER logFile
    Path to the log file where script actions and errors are recorded.

.EXAMPLE
    ./checkAccounts.ps1

    This command runs the script using the default file paths for CAPWATCH data and logs actions to `script_log.txt`.

.NOTES
    - Ensure the Microsoft Graph PowerShell SDK is installed and authenticated before running the script.
    - The script requires the following Microsoft Graph API permissions:
        - `User.Read.All`
        - `User.ReadWrite.All`
        - `Directory.ReadWrite.All`
    - The script assumes CAPID is stored in the `officeLocation` property of Azure AD users.
#>

# Input bindings are passed in via param block.
param($Timer)

# Set working directory to folder with all CAPWATCH CSV Text Files
$CAPWATCHDATADIR = "$($env:HOME)\data\CAPWatch"
Push-Location $CAPWATCHDATADIR

# Include shared Functions
. "$PSScriptRoot\..\shared\shared.ps1"
# Initialize dry-run mode
$dryRunMode = Get-DryRunMode
Write-Log "🔍 DRY-RUN MODE: $(if ($dryRunMode) { 'ENABLED (safe preview)' } else { 'DISABLED (EXECUTING CHANGES)' })"
#Abort script execution if CAPWATCH data is stale
$DownloadDate = (((Get-Date) - ((Import-Csv .\DownLoadDate.txt -ErrorAction Stop).DownLoadDate | Get-Date)).TotalHours)
Write-Log "Download date is: [$DownloadDate]"
if (((Get-Date) - ((Import-Csv .\DownLoadDate.txt -ErrorAction Stop).DownLoadDate | Get-Date)).TotalHours -gt 48) {
    Write-Error "CAPWATCH data in [$CAPWATCHDATADIR] is stale; aborting script execution!"
    exit 1
}

$MSGraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token

Connect-MgGraph -AccessToken $MSGraphAccessToken -NoWelcome
Connect-ExchangeOnline -ManagedIdentity -Organization $env:EXCHANGE_ORGANIZATION


# Import the CSV file into an array
$members = Import-Csv "$($CAPWATCHDATADIR)\Member.txt" -ErrorAction Stop | Where-Object { $_.MbrStatus -eq "ACTIVE" }
$dutyPositions_all = Import-Csv "$($CAPWATCHDATADIR)\DutyPosition.txt" -ErrorAction Stop
$contacts = Import-Csv "$($CAPWATCHDATADIR)\MbrContact.txt" -ErrorAction Stop

function Compare-Arrays {
    param (
        [array]$Array1,
        [array]$Array2
    )

    # Find user IDs that are in both arrays
    $inBoth = $Array1 | Where-Object { $Array2 -contains $_ }

    # Find user IDs that are only in Array1
    $AddtoTeams = $Array1 | Where-Object { $Array2 -notcontains $_ }

    # Find user IDs that are only in Array2
    $RemovefromTeams = $Array2 | Where-Object { $Array1 -notcontains $_ }

    # Output the results
    [PSCustomObject]@{
        InBoth       = $inBoth
        AddtoTeams = $AddtoTeams
        RemovefromTeams = $RemovefromTeams
    }
}

# This function combines the data from the Members and Contacts CSV files into a single array of objects.  This contains all information needed for O365.
# It creates a hashtable to store the combined data, where the key is the CAPID.

function Combine {
    param (
        [Array]$members,
        [Array]$contacts
    )

    # Initialize the hashtable to store combined data
    $combinedData = @{}
   # Add data from Members CSV to table
    foreach ($row in $members) {
#        Write-Log "Processing member: $($row.CAPID) - $($row.NameFirst) $($row.NameLast)"
        $combinedData[$row.CAPID] = @{
            CAPID = $row.CAPID
            NameLast = $row.NameLast
            NameFirst = $row.NameFirst
            Unit = $row.Unit
            Grade = $row.Rank
            Type = $row.Type
            Email = $null
            DoNotContact = $null
            MobilePhone = $null
            DOB = $row.DOB
            Joined = $row.Joined
        }
    }
# Add data from Contacts to table - Email and DoNotContact
foreach ($row in $contacts) {
    if ($combinedData.ContainsKey($row.CAPID)) {
        # Trim whitespace from contact data
        $contact = $row.Contact.Trim()
         # Assign EMAIL to cadet's Email field, but exclude CADET PARENT EMAIL (those get handled separately)
        if ($contact -match '^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$' -and $row.Priority -eq "PRIMARY" -and $row.Type -ne "CADET PARENT EMAIL") {
           $combinedData[$row.CAPID].Email = $contact
            $combinedData[$row.CAPID].DoNotContact = $row.DoNotContact
        }
        # Extract phone numbers - prefer CADET PARENT PHONE for mobilePhone, fallback to CELL PHONE
        if ($row.Priority -eq "PRIMARY") {
            if ($row.Type -eq "CADET PARENT PHONE" -and $row.Contact -match '^\+?[\d\-\(\)\s]+$') {
                $combinedData[$row.CAPID].MobilePhone = $row.Contact
            } elseif ($row.Type -eq "CELL PHONE" -and $row.Contact -match '^\+?[\d\-\(\)\s]+$' -and [string]::IsNullOrEmpty($combinedData[$row.CAPID].MobilePhone)) {
                $combinedData[$row.CAPID].MobilePhone = $row.Contact
            }
        }
        if ($row.Type -eq "CADET PARENT EMAIL" -and $contact -ne $combinedData[$row.CAPID].Email) {
            # Ensure the cadet entry exists before adding the parent
            if ($combinedData.ContainsKey($row.CAPID)) {
                $parentCAPID = "$($row.CAPID)P" # Use a unique key for the parent entry
                if (-not $combinedData.ContainsKey($parentCAPID)) {
                    $combinedData[$parentCAPID] = @{
                        CAPID = "$($row.CAPID)P"
                        NameLast = $combinedData[$row.CAPID].NameLast
                        NameFirst = $combinedData[$row.CAPID].NameFirst
                        Unit = $combinedData[$row.CAPID].Unit
                        Grade = "$($combinedData[$row.CAPID].Grade) PARENT"
                        Type = "PARENT"
                        Email = $contact
                        DoNotContact = $row.DoNotContact
                        MobilePhone = $combinedData[$row.CAPID].MobilePhone
                        DOB = $combinedData[$row.CAPID].DOB
                    }
                }
            } else {
                Write-Log "Warning: Parent email found for CAPID $($row.CAPID), but no cadet entry exists. Skipping parent entry."
            }
        }
    }
}
     # Convert the hashtable to an array
    $updates = $combinedData.Values

    $accountInfo = $updates | ForEach-Object {
        $obj = New-Object PSObject
        foreach ($key in $_.Keys) {
            $obj | Add-Member -MemberType NoteProperty -Name $key -Value $_[$key]
        }
        $obj    
    }

    $accountInfo 
}

# This function processes the Duty Positions CSV file and creates a hashtable where the key is the CAPID and the value is a string of duty positions.
# It also creates a string for each CAPID that contains the duty positions in the format "WING <positions> UNIT <positions>".
function DutyPositions {  
    param (
        [array]$dutyPositions_all
    )
    $capidPositions = @{}
    # Process each row in the CSV file
    foreach ($row in $dutyPositions_all) {
        $capid = $row.CAPID
        $functArea = $row.FunctArea
        $level = $row.Lvl

        if (-not [string]::IsNullOrEmpty($capid)) {
            if (-not $capidPositions.ContainsKey($capid)) {
                $capidPositions[$capid] = @{ 'WING' = @(); 'UNIT' = @() }
            }

            if ($level -eq 'WING' -or $level -eq 'UNIT') {
                $capidPositions[$capid][$level] += $functArea
            }
        }
    }
    $capidPositions
}

# This function processes the Duty Positions CSV file and creates a hashtable where the key is the CAPID and the value is a string of duty positions.
function MemberDuties {
    param (
        [array]$dutyPositions
    )

    # Initialize a hashtable to store positions for all CAPIDs
    $capidPositions = @{}

    # Process each row in the CSV file
    foreach ($row in $dutyPositions) {
        $capid = $row.CAPID
        $functArea = $row.FunctArea
        $level = $row.Lvl

        if (-not [string]::IsNullOrEmpty($capid)) {
            # Ensure the CAPID exists in the hashtable
            if (-not $capidPositions.ContainsKey($capid)) {
                $capidPositions[$capid] = @{ 'WING' = @(); 'UNIT' = @() }
            }

            # Add the duty position to the appropriate level (WING or UNIT)
            if ($level -eq 'WING' -or $level -eq 'UNIT') {
                if (-not ($capidPositions[$capid][$level] -contains $functArea)) {
                    $capidPositions[$capid][$level] += $functArea
                }
            }
        }
    }

    # Create an array to store the result for all CAPIDs
    $resultArray = @()

    foreach ($capid in $capidPositions.Keys) {
        # Remove duplicates and join positions for WING and UNIT
        $wingPositions = ($capidPositions[$capid]['WING'] | Sort-Object -Unique) -join ' '
        $unitPositions = ($capidPositions[$capid]['UNIT'] | Sort-Object -Unique) -join ' '

        # Construct the duty position string
        if ($wingPositions -ne '' -and $unitPositions -ne '') {
            $position = "WING $wingPositions UNIT $unitPositions"
        } elseif ($wingPositions -ne '') {
            $position = "WING $wingPositions"
        } elseif ($unitPositions -ne '') {
            $position = "UNIT $unitPositions"
        } else {
            $position = "No positions found for CAPID $capid"
        }

        # Add the CAPID and its positions to the result array
        $resultArray += [PSCustomObject]@{
            CAPID       = $capid
            DutyPosition = $position
        }
    }

    return $resultArray
}

# This function retrieves all users from Microsoft Graph API and returns them as an array.
function GetAllUsers {
    $allUsers = @()
    $uri = "https://graph.microsoft.com/beta/users?`$select=userPrincipalName,mail,displayName,officeLocation,companyName,employeeId,employeeType,jobTitle,department,mobilePhone,onPremisesExtensionAttributes,employeeHireDate"
    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
        # Flatten extension attributes for easier access
        foreach ($user in $response.value) {
            if ($user.onPremisesExtensionAttributes) {
                $user | Add-Member -NotePropertyName 'extensionAttribute1' -NotePropertyValue $user.onPremisesExtensionAttributes.extensionAttribute1 -Force
            }
        }
        $allUsers += $response.value
        $uri = $response.'@odata.nextLink'
    } while ($uri)
    return $allUsers
}

function GetGuestUserPrincipalName {
    param (
        [string]$email
    )

    if ([string]::IsNullOrWhiteSpace($email)) {
        return $null
    }

    $localPart = $email -replace '@', '_' -replace '[^a-zA-Z0-9._-]', ''
    return "$localPart#EXT#@$env:EXCHANGE_ORGANIZATION"
}

function GetGraphErrorDetailMessage {
    param (
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        return $ErrorRecord.ErrorDetails.Message
    }

    try {
        if ($ErrorRecord.Exception.Response -and $ErrorRecord.Exception.Response.Content) {
            $responseBody = $ErrorRecord.Exception.Response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            if (-not [string]::IsNullOrWhiteSpace($responseBody)) {
                return $responseBody
            }
        }
    } catch {
        # Best-effort only; fall back to the exception message below.
    }

    return $null
}

function GetConflictingExchangeRecipients {
    param (
        [string]$email
    )

    if ([string]::IsNullOrWhiteSpace($email)) {
        return @()
    }

    $smtpAddress = "SMTP:$email"

    try {
        return @(Get-Recipient -Filter "EmailAddresses -eq '$smtpAddress'" -ResultSize Unlimited -ErrorAction SilentlyContinue)
    } catch {
        Write-Log "Warning: Could not query Exchange recipients for $email. Error: $_"
        return @()
    }
}

function Resolve-ConflictingExchangeRecipients {
    param (
        [string]$email
    )

    $conflictingRecipients = @(GetConflictingExchangeRecipients -email $email)
    if ($conflictingRecipients.Count -eq 0) {
        return [PSCustomObject]@{
            Found        = $false
            Resolved     = $false
            RequiresManual = $false
        }
    }

    $removedAny = $false
    $requiresManual = $false

    foreach ($recipient in $conflictingRecipients) {
        $recipientType = [string]$recipient.RecipientTypeDetails
        $recipientIdentity = if ($recipient.Identity) { $recipient.Identity } else { $recipient.Name }
        $recipientDisplay = if ($recipient.DisplayName) { $recipient.DisplayName } else { $recipientIdentity }

        if ($recipientType -eq 'MailContact') {
            Write-OperationLog "Conflicting Exchange recipient found for $email" "$recipientDisplay [$recipientType] - removing MailContact"
            if (Test-ExecutionMode) {
                try {
                    Remove-MailContact -Identity $recipientIdentity -Confirm:$false -ErrorAction Stop
                    Write-Log "Deleted conflicting MailContact for ${email}: $recipientDisplay ($recipientIdentity)"
                    $removedAny = $true
                } catch {
                    Write-Log "Failed to delete conflicting MailContact for ${email}: $recipientDisplay ($recipientIdentity). Error: $_"
                    $requiresManual = $true
                }
            } else {
                Write-Log "[DRY-RUN] Would delete conflicting MailContact for ${email}: $recipientDisplay ($recipientIdentity)"
                $removedAny = $true
            }
        } else {
            Write-Log "Conflicting Exchange recipient remains for ${email}: $recipientDisplay [$recipientType] ($recipientIdentity). Manual cleanup required before guest invitation."
            $requiresManual = $true
        }
    }

    if ($removedAny -and (Test-ExecutionMode)) {
        Start-Sleep -Seconds 2
    }

    return [PSCustomObject]@{
        Found          = $true
        Resolved       = $removedAny -and -not $requiresManual
        RequiresManual = $requiresManual
    }
}

function AddNewGuest {
    param (
        [PSCustomObject]$userInfo,
        [array]$allUsers,
        [array]$deletedUsers
    )

    # Validate email before proceeding
    if (-not $userInfo.Email -or $userInfo.Email -notmatch '^[\w\.\-]+@([\w\-]+\.)+[\w\-]{2,}$') {
        Write-Log "Skipping guest creation: Missing or invalid email for CAPID $($userInfo.CAPID), Name: $($userInfo.NameFirst) $($userInfo.NameLast)"
        return
    }

    $userPrincipalName = GetGuestUserPrincipalName -email $userInfo.Email

    # Check if a deleted user exists with this CAPID or Email
    $restoreUser = $deletedUsers | Where-Object {
        $_.officeLocation -eq $userInfo.CAPID -or
        $_.mail -eq $userInfo.Email -or
        ($userPrincipalName -and $_.userPrincipalName -eq $userPrincipalName)
    } | Select-Object -First 1
    if ($restoreUser) {
        Write-Log "Deleted account found for CAPID: $($userInfo.CAPID), Email: $($restoreUser.displayName). Attempting to restore..."
        try {
            $restoreUri = "https://graph.microsoft.com/beta/directory/deletedItems/$($restoreUser.id)/restore"
            $restoredAccount = Invoke-MgGraphRequest -Method POST -Uri $restoreUri
            Write-Log "Successfully restored account: $($restoredAccount.displayName), Email: $($restoredAccount.mail)."
        } catch {
            Write-Log "Failed to restore deleted account for $($userInfo.Email). Error: $_"
        }
        return
    }

    # Skip users with verified domain emails (e.g., @cowg.cap.gov) - they should be regular members, not B2B guests
    if ($userInfo.Email -match '@cowg\.cap\.gov$') {
        Write-Log "Skipping guest creation for CAPID $($userInfo.CAPID): Email domain @cowg.cap.gov is a verified domain. User should be a regular member account."
        return
    }

    Write-Log "Adding guest $($userInfo.NameFirst) $($userInfo.NameLast), $($userInfo.Grade), $($userInfo.CAPID), $($userInfo.Email), $($env:WING_DESIGNATOR)-$($userInfo.Unit)"

    $existingUser = $null
    # Check if the userPrincipalName already exists in $allUsers
    $existingUser = $allUsers | Where-Object { $_.userPrincipalName -eq $userPrincipalName }

    if ($existingUser) {
        Write-Log "Skipping creation: User with userPrincipalName $userPrincipalName already exists in Azure AD. $($existingUser.id), $($existingUser.officeLocation), $($existingUser.displayName)"
        return
    }

    # CRITICAL CHECK: Verify email address is not already in use by any existing account
    # This prevents duplicate guest account creation attempts for the same email
    $existingWithEmail = $allUsers | Where-Object { $_.mail -eq $userInfo.Email } | Select-Object -First 1
    if ($existingWithEmail) {
        Write-Log "⚠️ [DUPLICATE EMAIL PREVENTION] Skipping account creation for CAPID $($userInfo.CAPID) ($($userInfo.NameFirst) $($userInfo.NameLast), $($userInfo.Grade)): Email $($userInfo.Email) already exists in Azure AD. Existing account: $($existingWithEmail.displayName) (CAPID: $($existingWithEmail.officeLocation), Type: $($existingWithEmail.employeeType)). This duplicate creation attempt has been logged and skipped."
        return
    }

    # Check for conflicting Exchange recipients with the same proxy address before inviting the guest.
    try {
        $conflictResolution = Resolve-ConflictingExchangeRecipients -email $userInfo.Email
        if ($conflictResolution.RequiresManual) {
            Write-Log "Skipping guest creation for $($userInfo.Email): A conflicting Exchange recipient still exists with this proxy address. Manual cleanup is required before inviting this user."
            return
        }
    } catch {
        Write-Log "Warning: Could not check for conflicting Exchange recipients for $($userInfo.Email): $_"
    }

    # Skip invitation for domains known to have cross-tenant restrictions
    if ($userInfo.Email -match '\.(mil|gov\.mil)$') {
        Write-Log "Skipping guest creation for $($userInfo.Email): Military (.mil) domains are typically blocked by cross-tenant access policies. Manual invitation required."
        return
    }

    # Use B2B invitation API to create guest user and send invitation in one step
    $invitationBody = @{
        invitedUserEmailAddress = $userInfo.Email
        inviteRedirectUrl = "https://myapplications.microsoft.com/?tenantid=71f5b48f-029d-4189-8fe0-052e14cec0ad"
        sendInvitationMessage = $true
        invitedUserDisplayName = "$($userInfo.NameFirst) $($userInfo.NameLast), $($userInfo.Grade)"
        invitedUserMessageInfo = @{
            customizedMessageBody = "Welcome to the Colorado Wing!`n`nYou are receiving this invitation because you or your cadet participates in Civil Air Patrol activities.`n`nWe are creating a secure Microsoft guest account for you so that we can deliver official CAP communication—such as squadron announcements, activity updates, and Wing-level notifications—in a reliable and secure way.`n`nBy default, this guest account only provides access to basic communication resources.`n`nIf your role ever requires access to additional CAP systems, those permissions would be granted separately and only when appropriate.`n`nPlease click the link below to accept the invitation and get started."
        }
    } | ConvertTo-Json -Depth 5

    try {
        # Create guest user via B2B invitation (creates account AND sends invitation)
        Write-OperationLog "Creating guest user via B2B invitation" "$($userInfo.Email) - $($userInfo.NameFirst) $($userInfo.NameLast), $($userInfo.Grade)"
        
        if (Test-ExecutionMode) {
            $invitationUri = "https://graph.microsoft.com/v1.0/invitations"
            $result = Invoke-MgGraphRequest -Method POST -Uri $invitationUri -Body $invitationBody -ContentType "application/json"
            
            $createdUserId = $result.invitedUser.id
            Write-Log "Guest user created successfully via B2B invitation: $($userInfo.Email), $($result.invitedUser.userPrincipalName), $createdUserId"
            Write-Log "B2B invitation sent successfully. Redemption URL: $($result.inviteRedeemUrl)"
        } else {
            Write-Log "[DRY-RUN] Guest user creation skipped (would create and send invitation)"
            return
        }
        
        # Update the created user with additional CAPID and unit information
        try {
            # Convert Joined date from MM/DD/YYYY to ISO 8601 format (YYYY-MM-DD)
            $isoJoinedDate = $null
            if ($userInfo.Joined) {
                try {
                    # Try multiple date formats to handle both M/dd/yyyy and MM/dd/yyyy
                    $parsedDate = $null
                    $dateFormats = @("M/dd/yyyy", "MM/dd/yyyy", "M/d/yyyy", "MM/dd/yy", "M/dd/yy")
                    foreach ($format in $dateFormats) {
                        try {
                            $parsedDate = [DateTime]::ParseExact($userInfo.Joined, $format, $null)
                            break
                        } catch {}
                    }
                    
                    if ($parsedDate) {
                        $isoJoinedDate = $parsedDate.ToString("yyyy-MM-dd")
                    } else {
                        Write-Log "Warning: Could not parse joined date for CAPID $($userInfo.CAPID): $($userInfo.Joined)"
                    }
                } catch {
                    Write-Log "Warning: Could not parse joined date for CAPID $($userInfo.CAPID): $($userInfo.Joined)"
                }
            }
            
            # Convert DOB date from MM/DD/YYYY to ISO 8601 format (YYYY-MM-DD)
            $isoDOB = $null
            if ($userInfo.DOB) {
                try {
                    # Try multiple date formats to handle both M/dd/yyyy and MM/dd/yyyy
                    $dobParsed = $null
                    $dateFormats = @("M/dd/yyyy", "MM/dd/yyyy", "M/d/yyyy", "MM/dd/yy", "M/dd/yy")
                    foreach ($format in $dateFormats) {
                        try {
                            $dobParsed = [DateTime]::ParseExact($userInfo.DOB, $format, $null)
                            break
                        } catch {}
                    }
                    
                    if ($dobParsed) {
                        $isoDOB = $dobParsed.ToString("yyyy-MM-dd")
                    } else {
                        Write-Log "Warning: Could not parse DOB for CAPID $($userInfo.CAPID): $($userInfo.DOB)"
                    }
                } catch {
                    Write-Log "Warning: Could not parse DOB for CAPID $($userInfo.CAPID): $($userInfo.DOB)"
                }
            }
            
            $updateBody = @{
                companyName = "CO-$($userInfo.Unit)"
                officeLocation = $userInfo.CAPID
                employeeId = $userInfo.CAPID
                jobTitle = $userInfo.Grade
                employeeType = $userInfo.Type
                mail = $userInfo.Email
                mobilePhone = $userInfo.MobilePhone
                employeeHireDate = $isoJoinedDate
                onPremisesExtensionAttributes = @{
                    extensionAttribute1 = $isoDOB
                }
            } | ConvertTo-Json -Depth 3
            
            $updateUri = "https://graph.microsoft.com/beta/users/$createdUserId"
            if (Test-ExecutionMode) {
                Invoke-MgGraphRequest -Method PATCH -Uri $updateUri -Body $updateBody -ContentType "application/json"
                Write-Log "Updated guest user with CAPID and unit information: $($userInfo.CAPID), CO-$($userInfo.Unit)"
            } else {
                Write-Log "[DRY-RUN] Would update guest user metadata (CAPID, unit, dates)"
            }
        } catch {
            Write-Log "Failed to update guest user metadata for $($userInfo.Email). Error: $_ (Invitation was still sent successfully)"
        }
        
        # Temporary: suppress welcome notification while doing mass account creation
        # To re-enable welcome emails, set $SEND_WELCOME_EMAIL = $true and uncomment the block below.
        $SEND_WELCOME_EMAIL = $true
        if ($SEND_WELCOME_EMAIL) {
            # Send notification email to recruiting distribution group for the unit
            $unitEmails = Get-UnitNotificationEmails -unit $userInfo.Unit
            Write-Log "This new user notification was also emailed to Unit: $unitEmails"
            # Send notification using Microsoft Graph API (recommended replacement for Send-MailMessage)
            try {
                $userPrincipalName = "cowg_it_helpdesk@cowg.cap.gov" # Use a service account or shared mailbox with Mail.Send permission
                # Build the toRecipients array
                $toRecipients = @(
                    @{ emailAddress = @{ address = "mike.schulte@cowg.cap.gov" } }
                )
                # Add the new user's email if not already present
                if ($userInfo.Email -and $userInfo.Email -ne "mike.schulte@cowg.cap.gov") {
                    $toRecipients += @{ emailAddress = @{ address = $userInfo.Email } }
                }
                foreach ($unitEmail in $unitEmails) {
                    if ($unitEmail -and $unitEmail -ne "mike.schulte@cowg.cap.gov" -and $unitEmail -ne $userInfo.Email) {
                        $toRecipients += @{ emailAddress = @{ address = $unitEmail } }
                    }
                }
                $mailBody = @{
                    message = @{
                        subject = "Welcome $($userInfo.Grade) $($userInfo.NameFirst) $($userInfo.NameLast) to CO-$($userInfo.Unit)"
                        body = @{
                            contentType = "HTML"
                            content = @"
<html>
  <body style='font-family: Arial, sans-serif; color: #222;'>
    <div style='text-align: center; margin-bottom: 20px;'>
      <img src='https://cowg.cap.gov/media/websites/COWG_T_7665FADF8B38C.PNG' alt='COWG Logo' style='max-width: 200px;'/>
    </div>
    <h2 style='color: #003366;'>Welcome $($userInfo.Grade) $($userInfo.NameFirst) $($userInfo.NameLast) to the Squadron!</h2>
    <p>Their COWG Guest account has been <b>created</b> and they will now receive COWG announcements and squadron emails.</p>
    <table style='margin: 20px auto; border-collapse: collapse;'>
      <tr><td style='padding: 4px 8px; font-weight: bold;'>Name:</td><td style='padding: 4px 8px;'>$($userInfo.NameFirst) $($userInfo.NameLast)</td></tr>
      <tr><td style='padding: 4px 8px; font-weight: bold;'>Grade:</td><td style='padding: 4px 8px;'>$($userInfo.Grade)</td></tr>
      <tr><td style='padding: 4px 8px; font-weight: bold;'>CAPID:</td><td style='padding: 4px 8px;'>$($userInfo.CAPID)</td></tr>
      <tr><td style='padding: 4px 8px; font-weight: bold;'>Email:</td><td style='padding: 4px 8px;'>$($userInfo.Email)</td></tr>
      <tr><td style='padding: 4px 8px; font-weight: bold;'>Unit:</td><td style='padding: 4px 8px;'>CO-$($userInfo.Unit)</td></tr>
    </table>
    <p style='font-size: 0.9em; color: #888; margin-top: 30px;'>This is an automated notification from the COWG IT Team.</p>
  </body>
</html>
"@
                        }
                        toRecipients = $toRecipients
                    }
                    saveToSentItems = $false
                } | ConvertTo-Json -Depth 4
                if (Test-ExecutionMode) {
                    $uri = "https://graph.microsoft.com/v1.0/users/$userPrincipalName/sendMail"
                    Invoke-MgGraphRequest -Method POST -Uri $uri -Body $mailBody -ContentType "application/json"
                    Write-Log "Notification email sent to mike.schulte@cowg.cap.gov via Microsoft Graph."
                } else {
                    Write-Log "[DRY-RUN] Would send welcome notification email"
                }
            } catch {
                Write-Log "Failed to send notification email via Microsoft Graph: $_"
            }
        }
    } catch {
        $errorMessage = $_.Exception.Message
        $graphErrorDetails = GetGraphErrorDetailMessage -ErrorRecord $_
        $fullErrorText = @($errorMessage, $graphErrorDetails) -join " "

        # Check for specific error scenarios and provide helpful messages
        if ($fullErrorText -match "conflicting contact object" -or $fullErrorText -match "same proxy address") {
            Write-Log "Failed to create guest user $($userInfo.Email): A conflicting Exchange recipient exists with this proxy address. Remove the conflicting contact/mail user in Exchange and retry."
        } elseif ($fullErrorText -match "cross-tenant access settings" -or $fullErrorText -match "blocked by cross-tenant") {
            Write-Log "Failed to create guest user $($userInfo.Email): Blocked by cross-tenant access settings. This domain may require admin approval or is restricted (common for .gov/.mil domains)."
        } elseif ($fullErrorText -match "invitation is not allowed") {
            Write-Log "Failed to create guest user $($userInfo.Email): B2B invitation not allowed for this domain. Check Azure AD External Identities settings."
        } else {
            if ($graphErrorDetails) {
                Write-Log "Failed to create guest user: $($userInfo.Email). Error: $errorMessage Details: $graphErrorDetails"
            } else {
                Write-Log "Failed to create guest user: $($userInfo.Email). Error: $errorMessage"
            }
        }
    }
}

# Helper function to get the recruiting distribution group email for a unit
function Get-UnitNotificationEmails {
    param (
        [string]$unit
    )
    $recruitingGroupEmail = "co-$($unit)-recruiting@cowg.cap.gov"
    return @($recruitingGroupEmail)
}

function AddNewAEMContact {
    param (
        [PSCustomObject]$userInfo
    )

    # Define the email to check
    $email = $userInfo.Email
    # Query for existing contact with this email
    $uri = "https://graph.microsoft.com/v1.0/contacts?`$filter=mail eq '$email'"
    $response = Invoke-MgGraphRequest -Method GET -Uri $uri

    if ($response.value.Count -gt 0) {
        Write-Log "Contact with email $email already exists. Skipping creation."
    } else {
        # Proceed to create the contact
        Write-OperationLog "Creating AEM contact" "$email - $($userInfo.NameFirst) $($userInfo.NameLast)"
        
        if (Test-ExecutionMode) {
            $contactBody = @{
                displayName = "$($userInfo.NameFirst) $($userInfo.NameLast), $($userInfo.Grade)"
                mailNickname = "$($userInfo.Email).Split('@')[0]"
                mail = "$($userInfo.Email)"
                userPrincipalName = "$($userInfo.Email)"
                givenName = "$($userInfo.NameFirst)"
                surname = "$($userInfo.NameLast)"
                companyName = "$($userInfo.Unit)"
                department = "AEM"
            } | ConvertTo-Json

            $createUri = "https://graph.microsoft.com/v1.0/contacts"
            Invoke-MgGraphRequest -Method POST -Uri $createUri -Body $contactBody -ContentType "application/json"
            Write-Log "Contact created: $email"
            # Send notification email
            Send-MailMessage -To 'mike.schulte@cowg.cap.gov' -From 'noreply@cowg.cap.gov' -Subject "New AEM Contact Added: $($userInfo.NameFirst) $($userInfo.NameLast)" -Body "A new AEM contact was added: $($userInfo.NameFirst) $($userInfo.NameLast), Grade: $($userInfo.Grade), CAPID: $($userInfo.CAPID), Email: $($userInfo.Email), Unit: CO-$($userInfo.Unit)" -SmtpServer 'smtp.office365.com' -UseSsl -Port 587
        } else {
            Write-Log "[DRY-RUN] Would create AEM contact and send notification"
        }
    }
}

function EnsureGuestMailProperty {
    param (
        [array]$allUsers,
        [array]$memberInfo
    )
    foreach ($user in $allUsers) {
        if ($user.userType -eq "Guest" -and ([string]::IsNullOrEmpty($user.mail))) {
            # Try to find the matching member by UPN or officeLocation
            $matchedMember = $memberInfo | Where-Object {
                ($_.CAPID -eq $user.officeLocation) -or
                ($user.userPrincipalName -like ("$($_.Email -replace '@', '_')#EXT#@$env:EXCHANGE_ORGANIZATION"))
            } | Select-Object -First 1

            if ($matchedMember -and $matchedMember.Email -and $matchedMember.Email -match '^[\w\.\-]+@([\w\-]+\.)+[\w\-]{2,}$') {
                # Check for mail/proxyAddresses conflict before attempting update
                $conflict = $allUsers | Where-Object {
                    ($_.mail -eq $matchedMember.Email -or ($_.proxyAddresses -contains ("SMTP:" + $matchedMember.Email))) -and $_.id -ne $user.id
                }
                if ($conflict) {
                    Write-Log "Skipped updating mail for $($user.displayName): email $($matchedMember.Email) already in use by another object."
                    continue
                }
                Write-OperationLog "Updating guest mail property" "$($user.displayName): $($matchedMember.Email)"
                if (Test-ExecutionMode) {
                    try {
                        $updateUri = "https://graph.microsoft.com/beta/users/$($user.id)"
                        $body = @{ mail = $matchedMember.Email } | ConvertTo-Json
                        Invoke-MgGraphRequest -Method PATCH -Uri $updateUri -Body $body -ContentType "application/json"
                    } catch {
                        # Fallback: log any other error
                        Write-Log "Failed to update mail property for $($user.displayName): $_"
                    }
                } else {
                    Write-Log "[DRY-RUN] Would update mail property for $($user.displayName)"
                }
            }
        }
    }
}

#see which users are missing and which users need to be deleted.
$bothUser = @()
$addUser = @()
$addMemberInfo = @()
$memberInfo = Combine -members $members -contacts $contacts
Write-Log "Number of members in combined data: $($memberInfo.Count)"
$dutyPositions = MemberDuties -dutyPositions $dutyPositions_all
$allUsers = GetAllUsers
$deletedDirectoryUsers = GetDeletedUsers
# Write-Output $memberInfo
$filteredMembers = $memberInfo | Where-Object { $_.Unit -ne "999" -and $_.Unit -ne "000" -and $_.DoNotContact -ne "True" -and $_.DoNotContact -ne $null -and $_.Type -ne "AEM" -and $_.Type -ne "PATRON" -and $_.MbrStatus -ne "EXPIRED" -and -not ($_.Email -and $_.Email -match '(?i)@coloradomilitaryacademy\.org$')}
$filteredMembers = $filteredMembers | Sort-Object -Property CAPID

# Deduplicate by email address - prefer SENIOR accounts, then non-PARENT versions (CAPID without "P" suffix)
$deduplicatedMembers = @{}
foreach ($member in $filteredMembers) {
    if ($member.Email) {
        $emailKey = $member.Email.ToLower()
        if (-not $deduplicatedMembers.ContainsKey($emailKey)) {
            $deduplicatedMembers[$emailKey] = $member
        } else {
            # Priority 1: Prefer SENIOR accounts
            $existingMember = $deduplicatedMembers[$emailKey]
            $currentIsSenior = $member.Type -match '(?i)SENIOR'
            $existingIsSenior = $existingMember.Type -match '(?i)SENIOR'
            
            # If current is SENIOR but existing is not, replace it
            if ($currentIsSenior -and -not $existingIsSenior) {
                Write-Log "🔄 [DEDUPLICATION] Email $($member.Email) found for both SENIOR (CAPID $($member.CAPID)) and non-SENIOR (CAPID $($existingMember.CAPID)) accounts. Keeping SENIOR account CAPID $($member.CAPID), removing non-SENIOR account CAPID $($existingMember.CAPID)."
                $deduplicatedMembers[$emailKey] = $member
            } elseif (-not $currentIsSenior -and $existingIsSenior) {
                # Existing is SENIOR, current is not - keep existing
                Write-Log "⚠️ [DUPLICATE EMAIL REMOVAL] Email $($member.Email) duplicated for CAPIDs $($member.CAPID) (non-SENIOR) and $($existingMember.CAPID) (SENIOR). Keeping SENIOR account (CAPID: $($existingMember.CAPID)), removing non-SENIOR duplicate (CAPID: $($member.CAPID))."
            } else {
                # Both are SENIOR or both are non-SENIOR, so apply Priority 2: prefer non-PARENT accounts
                $currentIsParent = $member.CAPID -match 'P$'
                $existingIsParent = $existingMember.CAPID -match 'P$'
                
                # If current is not parent but existing is parent, replace it
                if (-not $currentIsParent -and $existingIsParent) {
                    Write-Log "🔄 [DEDUPLICATION] Email $($member.Email) found in both member and parent accounts. Keeping non-PARENT account CAPID $($member.CAPID), removing parent account CAPID $($existingMember.CAPID)."
                    $deduplicatedMembers[$emailKey] = $member
                } else {
                    # Otherwise keep the existing one and log the duplicate
                    Write-Log "⚠️ [DUPLICATE EMAIL REMOVAL] Email $($member.Email) duplicated for CAPIDs $($member.CAPID) and $($existingMember.CAPID). Keeping first occurrence (CAPID: $($existingMember.CAPID)), removing duplicate."
                }
            }
        }
    } else {
        # Members without email - add using CAPID as temporary key to handle no-email cases
        if (-not $deduplicatedMembers.ContainsKey($member.CAPID)) {
            $deduplicatedMembers[$member.CAPID] = $member
        }
    }
}
$filteredMembers = $deduplicatedMembers.Values | Sort-Object -Property CAPID

if ($filteredMembers.Count -eq 0) {
    Write-Log "No filtered members found. Exiting the script."
    exit
}
Write-Log "filteredMembers: $($filteredMembers.count)"
$filteredMembers | Export-Csv -Path "$CAPWATCHDATADIR/FilteredMemberData.csv" -NoTypeInformation
Write-Log "Moving to member loop"
# Create a hash table for quick lookups of allUsers by officeLocation (CAPID)

# Normalize and create hash table for allUsers
$allUsersHash = @{}
foreach ($user in $allUsers) {
    if ($null -ne $user.officeLocation) {
        $normalizedOfficeLocation = $user.officeLocation
        $allUsersHash[$normalizedOfficeLocation] = $user
    }
}

# Initialize hash sets to avoid duplicates
$bothUserSet = @{}
$addUserSet = @{}

# Create hash table for mail lookups to avoid O(n) searches
$allUsersByMail = @{}
foreach ($user in $allUsers) {
    if ($user.mail) {
        $allUsersByMail[$user.mail.ToLower()] = $user
    }
}

# Process filteredMembers
foreach ($member in $filteredMembers) {
    # Check if the CAPID exists in the hash table
    $capidExists = $allUsersHash.ContainsKey($member.CAPID)
    
    # Also check by mail address using hash table for O(1) lookup
    $mailExists = if ($member.Email) { $allUsersByMail[$member.Email.ToLower()] } else { $null }
    
    if ($capidExists -or $mailExists) {
        if (-not $bothUserSet.ContainsKey($member.CAPID)) {
            $bothUser += $member.CAPID
            $bothUserSet[$member.CAPID] = $true
        }
    } else {
        # Only log new users being added, not expected matches
        if (-not $addUserSet.ContainsKey($member.CAPID)) {
            $addUser += $member.CAPID
            $addMemberInfo += $member
            $addUserSet[$member.CAPID] = $true
        }
    }
}
Write-Log "Add User count: $($addUser.Count)"

# Process members who do not currently have an active directory account
foreach ($user in $addUser) {
    $userInfo = $addMemberInfo | Where-Object { $_.CAPID -eq $user }
    if ($userInfo) {
        $guestUserPrincipalName = GetGuestUserPrincipalName -email $userInfo.Email
        # Check if the user needs to be restored (because they renewed their membership)
        $restoreUser = $deletedDirectoryUsers | Where-Object {
            $_.officeLocation -eq $userInfo.CAPID -or
            $_.mail -eq $userInfo.Email -or
            ($guestUserPrincipalName -and $_.userPrincipalName -eq $guestUserPrincipalName)
        } | Select-Object -First 1
        # Check if the email or CAPID already exists in $allUsers (by mail address or officeLocation)
        $existingUser = $allUsers | Where-Object { $_.mail -eq $userInfo.Email -or $_.officeLocation -eq $userInfo.CAPID } | Select-Object -First 1
        
        if ($restoreUser) {
            Write-OperationLog "Restoring deleted account" "CAPID: $($userInfo.CAPID) - $($restoreUser.displayName)"
            if (Test-ExecutionMode) {
                try {
                    # Restore the deleted account
                    $restoreUri = "https://graph.microsoft.com/beta/directory/deletedItems/$($restoreUser.id)/restore"
                    $restoredAccount = Invoke-MgGraphRequest -Method POST -Uri $restoreUri
                    Write-Log "Successfully restored account: $($restoredAccount.displayName), Email: $($restoredAccount.mail)."
                } catch {
                    Write-Log "Failed to restore deleted account for $($userInfo.Email). Error: $_"
                }
            } else {
                Write-Log "[DRY-RUN] Would restore deleted account: $($restoreUser.displayName)"
            }
        } elseif ($existingUser) {
            Write-Log "Skipping creation: User with email $($userInfo.Email) already exists in Azure AD. $($existingUser.id) $($existingUser.CAPID) $($existingUser.NameFirst) $($existingUser.NameLast)"
            continue
        } elseif ($userInfo.Type -eq 'AEM') { # Member is an AEM and should be added as a contact in Exchange
            Write-Log "Adding AEM $($userInfo.NameFirst) $($userInfo.NameLast), $($userInfo.Grade), $($userInfo.CAPID), $($userInfo.Email), CO-$($userInfo.Unit)"
            AddNewAEMContact -userInfo $userInfo
        } else {
            AddNewGuest -userInfo $userInfo -allUsers $allUsers -deletedUsers $deletedDirectoryUsers
        }
    }
}

# Ensure all guest users have the mail property set if possible
EnsureGuestMailProperty -allUsers $allUsers -memberInfo $memberInfo

# Create Duty Position Hash Table
### Below here sets the Department with all the Duty Positions for each member ###
$capidPositions = @{}
# Process each row in the CSV file
foreach ($row in $dutyPositions_all) {
    $capid = $row.CAPID
    $functArea = $row.FunctArea
    $level = $row.Lvl

    if (-not [string]::IsNullOrEmpty($capid)) {
        if (-not $capidPositions.ContainsKey($capid)) {
            $capidPositions[$capid] = @{ 'WING' = @(); 'UNIT' = @() }
        }

        if ($level -eq 'WING' -or $level -eq 'UNIT') {
            $capidPositions[$capid][$level] += $functArea
        }
    }
}

# Ensuring Correct CAPID, Duty Position, Type, and Unit Information
foreach ($contact in $filteredMembers) {
    $o365User = $allUsers | Where-Object { $contact.CAPID -eq $_.officeLocation } | Select-Object -First 1
    if ($o365User) {
        $updateNeeded = $false
        $updateReason = ""
        $updateParams = @{}

    if ($o365User.OfficeLocation -ne $contact.CAPID) {
        $updateParams["officeLocation"] = $contact.CAPID
        $updateNeeded = $true
        $updateReason += "OfficeLocation updated. "
    }

    if ($o365User.employeeID -ne $contact.CAPID) {
        $updateParams["employeeID"] = $contact.CAPID
        $updateNeeded = $true
        $updateReason += "EmployeeID updated. "
    }

    $unitNumber = "CO-$($contact.Unit)"
    if ($o365User.companyName -ne $unitNumber) {
        $updateParams["companyName"] = $unitNumber
        $updateNeeded = $true
        $updateReason += "CompanyName updated to $unitNumber. "
    }

    if ($o365User.employeeType -ne $contact.Type) {
        $updateParams["employeeType"] = $contact.Type
        $updateNeeded = $true
        $updateReason += "EmployeeType updated to $($contact.Type). "
    }

    if ($o365User.mail -ne $contact.Email) {
        if ($contact.Email -and $contact.Email -match '^[\w\.\-]+@([\w\-]+\.)+[\w\-]{2,}$' -and $o365User.mail -notmatch '@cowg\.cap\.gov$') {
            # Check for mail/proxyAddresses conflict before attempting update
            $conflict = $allUsers | Where-Object {
                ($_.mail -eq $contact.Email -or ($_.proxyAddresses -contains ("SMTP:" + $contact.Email))) -and $_.id -ne $o365User.id
            }
            if ($conflict) {
                Write-Log "Skipping mail update for $($contact.CAPID): Email $($contact.Email) already in use by another object: $($conflict.displayName), $($conflict.mail), $($conflict.officeLocation) (proxyAddresses conflict)."
            } else {
                $updateParams["mail"] = $contact.Email
                $updateNeeded = $true
                $updateReason += "Email updated to $($contact.Email) from $($o365User.mail). "
            }
        }
    }
    # New feature: If UPN ends with cowg.cap.gov but mail does not, set mail to cowg.cap.gov address
    if ($o365User.userPrincipalName -match '@cowg\.cap\.gov$' -and $o365User.mail -notmatch '@cowg\.cap\.gov$') {
        $cowgMail = $o365User.userPrincipalName
        if ($o365User.mail -ne $cowgMail) {
            # Check for mail/proxyAddresses conflict before attempting update
            $conflict = $allUsers | Where-Object {
                ($_.mail -eq $cowgMail -or ($_.proxyAddresses -contains ("SMTP:" + $cowgMail))) -and $_.id -ne $o365User.id
            }
            if ($conflict) {
                Write-Log "Skipping mail update for $($contact.CAPID): cowg.cap.gov mail $cowgMail already in use by another object: $($conflict.displayName), $($conflict.mail), $($conflict.officeLocation) (proxyAddresses conflict)."
            } else {
                $updateParams["mail"] = $cowgMail
                $updateNeeded = $true
                $updateReason += "Mail property set to cowg.cap.gov address $cowgMail based on UPN. "
            }
        }
    }

        # Get the duty positions for the current contact
        $memberDutyPosition = $dutyPositions | Where-Object { $_.CAPID -eq $contact.CAPID } | Select-Object -ExpandProperty DutyPosition
        if ($o365User.department -ne $memberDutyPosition) {
            $updateParams["department"] = $memberDutyPosition
            $updateNeeded = $true
            $updateReason += "Department updated to $memberDutyPosition. "
        }
        
        # Compare and update jobTitle (Grade)
        if ($o365User.jobTitle -ne $contact.Grade) {
            $updateParams["jobTitle"] = $contact.Grade
            $updateParams["displayName"] = "$($contact.NameFirst) $($contact.NameLast), $($contact.Grade)"
            $updateNeeded = $true
            $updateReason += "JobTitle updated to $($contact.Grade). DisplayName updated to $($contact.NameFirst) $($contact.NameLast), $($contact.Grade). "
        }

        # Compare and update mobilePhone (only if there's a value)
        if ($contact.MobilePhone -and $o365User.mobilePhone -ne $contact.MobilePhone) {
            $updateParams["mobilePhone"] = $contact.MobilePhone
            $updateNeeded = $true
            $updateReason += "MobilePhone updated to $($contact.MobilePhone). "
        }

        # Compare and update extensionAttribute1 (DOB)
        if ($contact.DOB) {
            # Validate DOB format - only update if it's a valid date
            try {
                # Try multiple date formats to handle both M/dd/yyyy and MM/dd/yyyy
                $dobParsed = $null
                $dateFormats = @("M/dd/yyyy", "MM/dd/yyyy", "M/d/yyyy", "MM/dd/yy", "M/dd/yy")
                foreach ($format in $dateFormats) {
                    try {
                        $dobParsed = [DateTime]::ParseExact($contact.DOB, $format, $null)
                        break
                    } catch {}
                }

                if ($dobParsed) {
                    $isoDOB = $dobParsed.ToString("yyyy-MM-dd")
                    # Normalize Azure value for comparison (trim whitespace, handle null)
                    $currentDOB = if ($o365User.extensionAttribute1) { $o365User.extensionAttribute1.ToString().Trim() } else { $null }
                    # Only update if the ISO format is different from what's in Azure
                    if ($currentDOB -ne $isoDOB) {
                        if (-not $updateParams.ContainsKey("onPremisesExtensionAttributes")) {
                            $updateParams["onPremisesExtensionAttributes"] = @{}
                        }
                        $updateParams["onPremisesExtensionAttributes"]["extensionAttribute1"] = $isoDOB
                        $updateNeeded = $true
                        $updateReason += "ExtensionAttribute1 (DOB) updated from '$currentDOB' to '$isoDOB'. "
                    }
                } else {
                    Write-Log "Warning: Could not parse DOB for CAPID $($contact.CAPID): $($contact.DOB). Skipping DOB update."
                }
            } catch {
                Write-Log "Warning: Could not parse DOB for CAPID $($contact.CAPID): $($contact.DOB). Skipping DOB update."
            }
        }

        # Compare and update employeeHireDate (Date Joined)
        # Convert date format from MM/DD/YYYY to ISO 8601 (YYYY-MM-DD) for Microsoft Graph compatibility
        if ($contact.Joined) {
            try {
                # Try multiple date formats to handle both M/dd/yyyy and MM/dd/yyyy
                $parsedDate = $null
                $dateFormats = @("M/dd/yyyy", "MM/dd/yyyy", "M/d/yyyy", "MM/dd/yy", "M/dd/yy")
                foreach ($format in $dateFormats) {
                    try {
                        $parsedDate = [DateTime]::ParseExact($contact.Joined, $format, $null)
                        break
                    } catch {}
                }

                if ($parsedDate) {
                    $isoDate = $parsedDate.ToString("yyyy-MM-dd")
                    # Normalize the Azure date for comparison (might have time component or different format)
                    $o365UserDate = $null
                    if ($o365User.employeeHireDate) {
                        try {
                            # Parse Azure's date and convert to same format for comparison
                            $azureDate = [DateTime]::Parse($o365User.employeeHireDate)
                            $o365UserDate = $azureDate.ToString("yyyy-MM-dd")
                        } catch {
                            $o365UserDate = $o365User.employeeHireDate
                        }
                    }
                    # Only update if dates are actually different
                    if ($isoDate -ne $o365UserDate) {
                        $updateParams["employeeHireDate"] = $isoDate
                        $updateNeeded = $true
                        $updateReason += "EmployeeHireDate (Date Joined) updated from '$o365UserDate' to '$isoDate'. "
                    }
                } else {
                    Write-Log "Warning: Could not parse joined date for CAPID $($contact.CAPID): $($contact.Joined). Skipping date update."
                }
            } catch {
                Write-Log "Warning: Could not parse date for CAPID $($contact.CAPID): $($contact.Joined). Skipping date update."
            }
        }

        if ($updateNeeded) {
            try {
                # Verify user still exists before attempting update
                $verifyUri = "https://graph.microsoft.com/beta/users/$($o365User.id)?`$select=id"
                try {
                    Invoke-MgGraphRequest -Method GET -Uri $verifyUri | Out-Null
                } catch {
                    # User no longer exists, skip silently
                    Write-Log "Skipping update: User $($o365User.id) no longer exists in Azure AD. CAPID: $($contact.CAPID), Email: $($contact.Email)"
                    continue
                }

                Write-OperationLog "Updating user in Entra ID" "$($contact.Email) - CAPID: $($contact.CAPID)"
                Write-Log "Update Reason: $updateReason"

                if (Test-ExecutionMode) {
                    $updateUri = "https://graph.microsoft.com/beta/users/$($o365User.id)"

                    # Remove any null values from updateParams before converting to JSON
                    $cleanParams = @{}
                    foreach ($key in $updateParams.Keys) {
                        if ($null -ne $updateParams[$key]) {
                            # Special handling for nested onPremisesExtensionAttributes
                            if ($key -eq "onPremisesExtensionAttributes" -and $updateParams[$key] -is [hashtable]) {
                                # Only include if it has non-null values inside
                                $nestedParams = @{}
                                foreach ($nestedKey in $updateParams[$key].Keys) {
                                    if ($null -ne $updateParams[$key][$nestedKey]) {
                                        $nestedParams[$nestedKey] = $updateParams[$key][$nestedKey]
                                    }
                                }
                                if ($nestedParams.Count -gt 0) {
                                    $cleanParams[$key] = $nestedParams
                                }
                            } else {
                                $cleanParams[$key] = $updateParams[$key]
                            }
                        }
                    }
                    $body = $cleanParams | ConvertTo-Json -Depth 3
                    Invoke-MgGraphRequest -Method PATCH -Uri $updateUri -Body $body -ContentType "application/json"
                    Write-Log "Updated user: $($contact.Email), CAPID: $($contact.CAPID), Unit: $($contact.Unit), Duty Position: $memberDutyPosition, $($contact.Type))"
                } else {
                    Write-Log "[DRY-RUN] Would update user: $($contact.Email) - $updateReason"
                }
            } catch {
                $errorMessage = $_.Exception.Message
                $fullError = $_
                # Check for actual HTTP 404 (user not found in Azure AD)
                if ($fullError.Exception.Response.StatusCode -eq 404 -or $errorMessage -match "^Not Found") {
                    Write-Log "Skipping update: User $($o365User.id) no longer exists in Azure AD or is inaccessible. CAPID: $($contact.CAPID), Email: $($contact.Email)"
                } else {
                    Write-Log "Failed to update user: $($contact.Email). Error: $errorMessage"
                }
            }
        }
    }
}

Write-Log "Number in Both"
Write-Log $bothUser.count
Write-Log "Need to add users to O365"
# $addMemberInfo | Export-Csv -Path "./addMemberInfo.csv" -NoTypeInformation
Write-Log $addUser.Count

#Users with no CAPID...
$noCAPID = $allUsers | Where-Object {$_.officeLocation -eq $null }
# Write-Output $noCAPID | Select-Object displayName, mail | Format-Table -AutoSize
$noCAPID | Select-Object displayName, mail | Export-Csv -Path "./noCAPID.csv" -NoTypeInformation

# Group users by displayName and filter groups with more than one user
$duplicateDisplayNames = $allUsers | Group-Object -Property displayName | Where-Object { $_.Count -gt 1 }

# Output duplicate display names and their associated accounts
if ($duplicateDisplayNames.Count -gt 0) {
    Write-Log "Accounts with duplicate display names:"
    $duplicateDisplayNames | ForEach-Object {
        Write-Log "Display Name: $($_.Name)"
        $_.Group | Select-Object displayName, mail, officeLocation | Format-Table -AutoSize
        Write-Log "----------------------------------------"
    }
} else {
    Write-Log "No duplicate display names found."
}

Write-Log "Account deletion for expired members has been moved to the Maintenance function and runs on the 3rd of each month."
