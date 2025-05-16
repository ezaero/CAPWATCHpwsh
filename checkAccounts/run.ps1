
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
$members = Import-Csv "$($CAPWATCHDATADIR)\Member.txt" -ErrorAction Stop
$dutyPositions_all = Import-Csv "$($CAPWATCHDATADIR)\DutyPosition.txt" -ErrorAction Stop
$contacts = Import-Csv "$($CAPWATCHDATADIR)\MbrContact.txt" -ErrorAction Stop

# This function compares two arrays and returns the user IDs that are in both, only in the first array, and only in the second array.
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
        [PSCustomObject]$userInfo
    )

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
        userPrincipalName = $userPrincipalName
        userType = "Guest"
        companyName = "CO-$($userInfo.Unit)" # Store the unit information
        officeLocation = $userInfo.CAPID # Store CAPID in officeLocation for easy lookup
        employeeId = $userInfo.CAPID # Store CAPID in department
        jobTitle = $userInfo.Grade
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
    } catch {
        Write-Log "Failed to create guest user: $($userInfo.Email). Error: $_"
    }
}

function RestoreDeletedAccounts {
    param (
        [Array]$expiredMembers
    )

    # Define the API endpoint to query deleted users
    $deletedUsersUri = "https://graph.microsoft.com/beta/directory/deletedItems/microsoft.graph.user"

    try {
        # Retrieve deleted users
        $deletedUsers = @()
        do {
            $response = Invoke-MgGraphRequest -Method GET -Uri $deletedUsersUri
            $deletedUsers += $response.value
            $deletedUsersUri = $response.'@odata.nextLink'
        } while ($deletedUsersUri)

        # Loop through each expired member
        foreach ($expiredMember in $expiredMembers) {
            $capid = $expiredMember.CAPID
            $email = $expiredMember.Email

            # Check if the CAPID or Email matches a deleted user
            $deletedAccount = $deletedUsers | Where-Object {
                $_.officeLocation -eq $capid -or $_.mail -eq $email
            }

            if ($deletedAccount) {
                Write-Log "Deleted account found for CAPID: $capid, Email: $email. Attempting to restore..."
                # Restore the deleted account
                $restoreUri = "https://graph.microsoft.com/beta/directory/deletedItems/$($deletedAccount.id)/restore"
                $restoredAccount = Invoke-MgGraphRequest -Method POST -Uri $restoreUri

                Write-Log "Successfully restored account: $($restoredAccount.displayName), Email: $($restoredAccount.mail)."
            } else {
                Write-Log "No deleted account found for CAPID: $capid, Email: $email."
           }
        }
    } catch {
        Write-Log "Failed to check or restore deleted accounts. Error: $_"
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
# $filteredMembers | Export-Csv -Path ../output/FilteredMemberData.csv -NoTypeInformation
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

            # Restore the deleted account
            $restoreUri = "https://graph.microsoft.com/beta/directory/deletedItems/$($restoreUser.id)/restore"
            $restoredAccount = Invoke-MgGraphRequest -Method POST -Uri $restoreUri

            Write-Log "Successfully restored account: $($restoredAccount.displayName), Email: $($restoredAccount.mail)."
        } elseif ($existingUser) {
            Write-Log "Skipping creation: User with email $($userInfo.Email) already exists in Azure AD. $($userInfo.id) $($userInfo.CAPID) $($userInfo.NameFirst) $($userInfo.NameLast)"
            continue
        } else {
            # Call AddNewGuest if the email does not exist
            AddNewGuest -userInfo $userInfo
        }
    }
}

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
        $updateParams = @{}

        if ($o365User.OfficeLocation -ne $contact.CAPID) {
            $updateParams["officeLocation"] = $contact.CAPID
            $updateNeeded = $true
        }

        if ($o365User.employeeID -ne $contact.CAPID) {
            $updateParams["employeeID"] = $contact.CAPID
            $updateNeeded = $true
        }

        $unitNumber = "CO-$($contact.Unit)"
        if ($o365User.companyName -ne $unitNumber) {
            $updateParams["companyName"] = $unitNumber
            $updateNeeded = $true
        }

        if ($o365User.employeeType -ne $contact.Type) {
            $updateParams["employeeType"] = $contact.Type
            $updateNeeded = $true
        }

        # Get the duty positions for the current contact
        $memberDutyPosition = $dutyPositions | Where-Object { $_.CAPID -eq $contact.CAPID } | Select-Object -ExpandProperty DutyPosition
        if ($o365User.department -ne $memberDutyPosition) {
            $updateParams["department"] = $memberDutyPosition
            $updateNeeded = $true
        }

        if ($updateNeeded) {
            try {
                $updateUri = "https://graph.microsoft.com/beta/users/$($o365User.id)"
                $body = $updateParams | ConvertTo-Json
                Invoke-MgGraphRequest -Method PATCH -Uri $updateUri -Body $body -ContentType "application/json"
                Write-Log "Updated user: $($o365User.mail), CAPID: $($contact.CAPID), Unit: $($contact.Unit), Duty Position: $memberDutyPosition, $($contact.Type))"
            } catch {
                Write-Log "Failed to update user: $($o365User.mail). Error: $_"
           }
        }
    }
}

Write-Log "Number in Both"
$bothUser.count
Write-Log "Need to add users to O365"
# $addMemberInfo | Export-Csv -Path "./addMemberInfo.csv" -NoTypeInformation
$addUser.Count

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

# # Loop through each expired member
# foreach ($expiredMember in $expiredMembers) {
#     $capid = $expiredMember.CAPID
#     $parentCAPID = "$capid`P" # Parent's CAPID is CAPID + "P"

#     # Find the member's account in Azure AD
#     $memberAccount = $allUsers | Where-Object { $_.officeLocation -eq $capid }
#     $parentAccount = $allUsers | Where-Object { $_.officeLocation -eq $parentCAPID }

#     # Delete the member's account
#     if ($memberAccount) {
#         try {
#             $uri = "https://graph.microsoft.com/v1.0/users/$($memberAccount.id)"
#             Invoke-MgGraphRequest -Method DELETE -Uri $uri
#             Write-Log "Deleted member account: $($memberAccount.displayName) ($($memberAccount.mail)) with CAPID: $capid."
#         } catch {
#             Write-Log "Failed to delete member account: $($memberAccount.displayName) ($($memberAccount.mail)). Error: $_"
#         }
#     } else {
# #        Write-Log "No member account found for CAPID: $capid."
#     }

#     # Delete the parent's guest account
#     if ($parentAccount) {
#         try {
#             $uri = "https://graph.microsoft.com/v1.0/users/$($parentAccount.id)"
#             Invoke-MgGraphRequest -Method DELETE -Uri $uri
#             Write-Log "Deleted parent account: $($parentAccount.displayName) ($($parentAccount.mail)) with CAPID: $parentCAPID."
#         } catch {
#             Write-Log "Failed to delete parent account: $($parentAccount.displayName) ($($parentAccount.mail)). Error: $_"
#         }
#     } else {
# #        Write-Log "No parent account found for CAPID: $parentCAPID."
#     }
# }


