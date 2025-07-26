<#
.SYNOPSIS
    Synchronizes CAPWATCH data with Microsoft Entra ID (Azure AD) and ensures accurate user information in O365.

.DESCRIPTION
    This script performs the following tasks:
    1. Connects to Microsoft Graph API using the Microsoft Graph PowerShell SDK.
    2. Imports data from CAPWATCH CSV files (`MbrContact.txt`, `Member.txt`, and `DutyPosition.txt`).
    3. Combines the data from the CSV files into a unified dataset for processing.
    4. Compares the CAPWATCH data with existing Microsoft Entra ID (Azure AD) users to:
        - Identify users to be added as O365 guest accounts.
        - Identify users to be removed from O365.
        - Ensure all O365 accounts have the correct CAPID, duty positions, and unit information.
    5. Creates O365 guest accounts for users missing in Azure AD.
    6. Updates existing O365 accounts with CAPID, duty positions, and unit information.
    7. Removes users from O365 who are no longer in the CAPWATCH data.
    8. Identifies and logs users with duplicate display names in Azure AD.
    9. Exports users with missing CAPIDs and logs all actions for auditing purposes.

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

#Abort script execution if CAPWATCH data is stale
$DownloadDate = (((Get-Date) - ((Import-Csv .\DownLoadDate.txt -ErrorAction Stop).DownLoadDate | Get-Date)).TotalHours)
Write-Log "Download date is: [$DownloadDate]"
if (((Get-Date) - ((Import-Csv .\DownLoadDate.txt -ErrorAction Stop).DownLoadDate | Get-Date)).TotalHours -gt 48) {
    Write-Error "CAPWATCH data in [$CAPWATCHDATADIR] is stale; aborting script execution!"
    exit 1
}

$MSGraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token

Connect-MgGraph -AccessToken $MSGraphAccessToken -NoWelcome
Connect-ExchangeOnline -ManagedIdentity -Organization COCivilAirPatrol.onmicrosoft.com


# Import the CSV file into an array
$members = Import-Csv "$($CAPWATCHDATADIR)\Member.txt" -ErrorAction Stop | Where-Object { $_.MbrStatus -eq "ACTIVE" }
$expiredMembers = Import-Csv "$($CAPWATCHDATADIR)\Member.txt" -ErrorAction Stop| Where-Object { $_.MbrStatus -eq "EXPIRED" }
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
        }
    }
# Add data from Contacts to table - Email and DoNotContact
foreach ($row in $contacts) {
    if ($combinedData.ContainsKey($row.CAPID)) {
        if ($row.Contact -match '^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$' -and $row.Priority -eq "PRIMARY") {
            $combinedData[$row.CAPID].Email = $row.Contact
            $combinedData[$row.CAPID].DoNotContact = $row.DoNotContact
        }
        if ($row.Type -eq "CADET PARENT EMAIL" -and $row.contact -ne $combinedData[$row.CAPID].Email) {
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
                        Email = $row.Contact
                        DoNotContact = $row.DoNotContact
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
    $uri = "https://graph.microsoft.com/beta/users?$select=mail,displayName,officeLocation,companyName,employeeId"
    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
        $allUsers += $response.value
        $uri = $response.'@odata.nextLink'
    } while ($uri)
    return $allUsers
}

function AddNewGuest {
    param (
        [PSCustomObject]$userInfo,
        [array]$allUsers
    )

    # Validate email before proceeding
    if (-not $userInfo.Email -or $userInfo.Email -notmatch '^[\w\.\-]+@([\w\-]+\.)+[\w\-]{2,}$') {
        Write-Log "Skipping guest creation: Missing or invalid email for CAPID $($userInfo.CAPID), Name: $($userInfo.NameFirst) $($userInfo.NameLast)"
        return
    }

    # Check if a deleted user exists with this CAPID or Email
    $restoreUser = $deletedUsers | Where-Object { $_.officeLocation -eq $userInfo.CAPID -or $_.mail -eq $userInfo.Email } | Select-Object -First 1
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

    Write-Log "Adding guest $($userInfo.NameFirst) $($userInfo.NameLast), $($userInfo.Grade), $($userInfo.CAPID), $($userInfo.Email), CO-$($userInfo.Unit)"
  
    # Replace '@' with '_' and remove invalid characters
    $localPart = $userInfo.Email -replace '@', '_' -replace '[^a-zA-Z0-9._-]', ''

    # Append '#EXT#' and the tenant domain
    $userPrincipalName = "$localPart#EXT#@COCivilAirPatrol.onmicrosoft.com"

    $existingUser = $null
    # Check if the userPrincipalName already exists in $allUsers
    $existingUser = $allUsers | Where-Object { $_.userPrincipalName -eq $userPrincipalName }

    if ($existingUser) {
        Write-Log "Skipping creation: User with userPrincipalName $userPrincipalName already exists in Azure AD. $($existingUser.id), $($existingUser.officeLocation), $($existingUser.displayName)"
        return
    }
    
    $body = @{
        accountEnabled = $true
        displayName = "$($userInfo.NameFirst) $($userInfo.NameLast), $($userInfo.Grade)"
        mailNickname = $($userInfo.Email).Split('@')[0] # Use the part before '@' as the mailNickname
        mail = $userInfo.Email # Always set to real email
        userPrincipalName = $userPrincipalName
        userType = "Guest"
        companyName = "CO-$($userInfo.Unit)" # Store the unit information
        officeLocation = $userInfo.CAPID # Store CAPID in officeLocation for easy lookup
        employeeId = $userInfo.CAPID # Store CAPID in department
        jobTitle = $userInfo.Grade
        employeeType = $userInfo.Type # CADET, PARENT, SENIOR, AEM, etc.
        passwordProfile = @{
            forceChangePasswordNextSignIn = $false
            password = "DummyPassword123!" # A dummy password to satisfy the API
        }
    } | ConvertTo-Json -Depth 2

    # Define the API endpoint
    $uri = "https://graph.microsoft.com/beta/users"

    try {
        # Create the guest user
        $result = Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -ContentType "application/json"
        Write-Log "Guest user created successfully: $($userInfo.Email), $($result.userPrincipalName), $($result.id)"
        # Send notification email to commanders and recruiting officer of the unit
        $unitEmails = Get-UnitNotificationEmails -unit $userInfo.Unit -allUsers $allUsers
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
            $uri = "https://graph.microsoft.com/v1.0/users/$userPrincipalName/sendMail"
            Invoke-MgGraphRequest -Method POST -Uri $uri -Body $mailBody -ContentType "application/json"
            Write-Log "Notification email sent to mike.schulte@cowg.cap.gov via Microsoft Graph."
        } catch {
            Write-Log "Failed to send notification email via Microsoft Graph: $_"
        }
    } catch {
        Write-Log "Failed to create guest user: $($userInfo.Email). Error: $_"
    }
}

# Helper function to get notification emails for a unit's commanders and recruiting officer
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
                ($user.userPrincipalName -like ("$($_.Email -replace '@', '_')#EXT#@COCivilAirPatrol.onmicrosoft.com"))
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
                Write-Log "Updating mail property for guest user $($user.displayName) ($($user.userPrincipalName)) to $($matchedMember.Email)"
                try {
                    $updateUri = "https://graph.microsoft.com/beta/users/$($user.id)"
                    $body = @{ mail = $matchedMember.Email } | ConvertTo-Json
                    Invoke-MgGraphRequest -Method PATCH -Uri $updateUri -Body $body -ContentType "application/json"
                } catch {
                    # Fallback: log any other error
                    Write-Log "Failed to update mail property for $($user.displayName): $_"
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
$deletedUsers = GetDeletedUsers
# Write-Output $memberInfo
$filteredMembers = $memberInfo | Where-Object { $_.Unit -ne "999" -and $_.Unit -ne "000" -and $_.DoNotContact -ne "True" -and $_.DoNotContact -ne $null -and $_.Type -ne "AEM" -and $_.Type -ne "PATRON" -and $_.MbrStatus -ne "EXPIRED" }
$filteredMembers = $filteredMembers | Sort-Object -Property CAPID
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

# Process filteredMembers
foreach ($member in $filteredMembers) {
    # Check if the CAPID or Email exists in the hash table
    $capidExists = $allUsersHash.ContainsKey($member.CAPID)
    # Replace '@' with '_' and remove invalid characters
    $localPart = $member.Email -replace '@', '_' -replace '[^a-zA-Z0-9._-]', ''
    # Append '#EXT#' and the tenant domain
    $userPrincipalName = "$localPart#EXT#@COCivilAirPatrol.onmicrosoft.com"
    $upnExists = $allUsers | Where-Object { $_.userPrincipalName -eq $userPrincipalName }

    if ($capidExists -or $upnExists) {
        if (-not $bothUserSet.ContainsKey($member.CAPID)) {
            $bothUser += $member.CAPID
            $bothUserSet[$member.CAPID] = $true
        }
    } else {
        Write-Log "CAPID $($member.CAPID) or UPN $userPrincipalName not found in allUsers."
        if (-not $addUserSet.ContainsKey($member.CAPID)) {
            $addUser += $member.CAPID
            $addMemberInfo += $member
            $addUserSet[$member.CAPID] = $true
        }
    }
}
Write-Log "Add User count: $($addUser.Count)"
foreach ($user in $addUser) {
    $userInfo = $addMemberInfo | Where-Object { $_.CAPID -eq $user }
    if ($userInfo) {
        # Check if the user needs to be restored (because they renewed their membership)
        $restoreUser = $deletedUsers | Where-Object { $_.officeLocation -eq $userInfo.CAPID } | Select-Object -First 1
        # Check if the email already exists in $allUsers
        $existingUser = $allUsers | Where-Object { $_.mail -eq $userInfo.Email -or $_.officeLocation -eq $userInfo.CAPID }
        if ($restoreUser) {
            Write-Log "Deleted account found for CAPID: $($userInfo.CAPID), Email: $($restoreUser.displayName). Attempting to restore..."
            try {
                # Restore the deleted account
                $restoreUri = "https://graph.microsoft.com/beta/directory/deletedItems/$($restoreUser.id)/restore"
                $restoredAccount = Invoke-MgGraphRequest -Method POST -Uri $restoreUri
                Write-Log "Successfully restored account: $($restoredAccount.displayName), Email: $($restoredAccount.mail)."
            } catch {
                Write-Log "Failed to restore deleted account for $($userInfo.Email). Error: $_"
            }
        } elseif ($existingUser) {
            Write-Log "Skipping creation: User with email $($userInfo.Email) already exists in Azure AD. $($userInfo.id) $($userInfo.CAPID) $($userInfo.NameFirst) $($userInfo.NameLast)"
            continue
        } elseif ($userInfo.Type -eq 'AEM') { # Member is an AEM and should be added as a contact in Exchange
            Write-Log "Adding AEM $($userInfo.NameFirst) $($userInfo.NameLast), $($userInfo.Grade), $($userInfo.CAPID), $($userInfo.Email), CO-$($userInfo.Unit)"
            AddNewAEMContact -userInfo $userInfo
        } else {
            AddNewGuest -userInfo $userInfo -allUsers $allUsers
        }
    }
}

# Ensure all guest users have the mail property set if possible
EnsureGuestMailProperty -allUsers $allUsers -memberInfo $memberInfo

# Create a list of CAPIDs from bothUser and addUser
$bothUserCAPIDs = $bothUser
$addUserCAPIDs = $addUser

# Select deletedUsers where the CAPID is not in bothUser or addUser
$deletedUsers = $filteredMembers | Where-Object { 
    -not ($_.CAPID -in $bothUserCAPIDs) -and -not ($_.CAPID -in $addUserCAPIDs) 
}

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

        if ($updateNeeded) {
            Write-Log "Attempting to update user: $($contact.Email), CAPID: $($contact.CAPID), Unit: $($contact.Unit), Duty Position: $memberDutyPosition, $($contact.Type))"
            Write-Log "Update Reason: $updateReason"
            try {
                $updateUri = "https://graph.microsoft.com/beta/users/$($o365User.id)"
                $body = $updateParams | ConvertTo-Json
                Invoke-MgGraphRequest -Method PATCH -Uri $updateUri -Body $body -ContentType "application/json"
                Write-Log "Updated user: $($contact.Email), CAPID: $($contact.CAPID), Unit: $($contact.Unit), Duty Position: $memberDutyPosition, $($contact.Type))"
            } catch {
                Write-Log "Failed to update user: $($contact.Email). Error: $_"
                # Check for any object with the same proxy address
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