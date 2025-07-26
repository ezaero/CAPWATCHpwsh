# Input bindings are passed in via param block.
param($Timer)

# Set working directory to folder with all CAPWATCH CSV Text Files
$CAPWATCHDATADIR = "$($env:HOME)\data\CAPWatch"
Push-Location $CAPWATCHDATADIR

$OrganizationFile = "$($CAPWATCHDATADIR)/Organization.txt"
$CommandersFile = "$($CAPWATCHDATADIR)/Commanders.txt"


#Abort script execution if CAPWATCH data is stale
if (((Get-Date) - ((Import-Csv .\DownLoadDate.txt -ErrorAction Stop).DownLoadDate | Get-Date)).TotalHours -gt 40) {
    Write-Error "CAPWATCH data in [$CAPWATCHDATADIR] is stale; aborting script execution!"
    exit 1
}

# Define the path to the DutyPosition.txt file
# Read the DutyPosition.txt file and reduce it to only WING Staff
$dutyPositions_all = Import-Csv .\DutyPosition.txt -ErrorAction Stop
$dutyPositions = $dutyPositions_all | Where-Object { $_.Lvl -eq "WING" }
$dutyPositions = $dutyPositions | Sort-Object CAPID -Unique

# Connect to Microsoft Graph
$MSGraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token
Connect-MgGraph -AccessToken $MSGraphAccessToken -NoWelcome
Connect-ExchangeOnline -ManagedIdentity -Organization COCivilAirPatrol.onmicrosoft.com


# This function takes 2 arrays and compares them, return 3 Arrays
function Compare-UserIds {
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
  
# Takes in an Array of userId's and prints the display names of each of them
function Get-DisplayNames {
    param (
        [array]$UserIds
    )

    foreach ($userId in $UserIds) {
        try {
            # Define the API endpoint
            $uri = "https://graph.microsoft.com/v1.0/users/$userId"
    
            # Make the API request to get the user details
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ContentType "application/json"
    
            # Extract and print the displayName
            $displayName = $response.displayName
            Write-Log "User ID: $userId, Display Name: $displayName"
        } catch {
            Write-Error "Failed to retrieve display name for User ID: $userId. Error: $_"
        }
    }
}

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

function GetAllUsers {
    # Fetch all users from Microsoft Entra ID (Azure AD)
    $allUsers = @()
    $uri = "https://graph.microsoft.com/beta/users?$select=id,displayName,officeLocation,companyName"
    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
        $allUsers += $response.value
        $uri = $response.'@odata.nextLink'
    } while ($uri)
    $allUsers
}

function GetAllGroups {
    # Get a list of groups that contain teams
    $allGroups = @()
    $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')"

    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
        $allGroups += $response.value
        $uri = $response.'@odata.nextLink'
    } while ($uri)
    $allGroups
}

function GetUnits {
    # Create a list of all Unit charter numbers and names in the Wing
    $organization_all = Import-Csv -Path $OrganizationFile
    $co_org = $organization_all | Where-Object { $_.Wing -eq "CO" } | Sort-Object Unit -Unique
    $co_org = $co_org | Select-Object Unit, Name
    # unitList will be a list of all the Teams required
    $unitList = @()
    foreach ($unit in $co_org) {
        if ($unit.Unit -ne "000" -and $unit.Unit -ne "999" -and $unit.Unit -ne "001") {
            $unitList += "CO-$($unit.Unit) $($unit.Name)"
        }
    }
    $unitList
}

function GetCommander {
    param (
        [Array]$allUsers,
        [string]$unit
    )

    $commanders_all = @()
    $commander = @()
    $commanders_all = Import-Csv -Path $CommandersFile
    
    $commander = $commanders_all | Where-Object { $_.Unit -eq $unit -and $_.Wing -eq 'CO' }

    $CAPID = $commander.CAPID
    Write-Host "Commander CAPID is: $CAPID"
    $commander
}

# Define the function to check if a team exists
function CheckTeamExists {
    param (
        [string]$teamName
    )

    # Define the endpoint to get the team by display name
    $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$teamName' and resourceProvisioningOptions/Any(x:x eq 'Team')"

    # Make the API request
    $response = Invoke-MgGraphRequest -Method GET -Uri $uri

    if ($response.value.Count -gt 0) {
        return $true
    } else {
        return $false
    }
}

# Function to generate a camel-cased alias for a team
function New-TeamAlias {
    param (
        [string]$unitName,
        [string]$domain = "cowg.cap.gov"  # Default domain
    )

    # Remove the unit designator (e.g., "CO-148")
    $squadronName = $unitName -replace "^[A-Z]{2}-\d{3}\s", ""

    # Convert to Camel Case
    $camelCasedName = ($squadronName -split '\s' | ForEach-Object {
        $_.Substring(0, 1).ToUpper() + $_.Substring(1).ToLower()
    }) -join ""

    # Append the domain
    $alias = "$camelCasedName@$domain"

    return $alias
}

function CheckTeams {
    param (
        [Array]$unitList,
        [Array]$allGroups,
        [Array]$allUsers
    )

    foreach ($unitName in $unitList) {
        $teamExists = $false
        $teamExists = CheckTeamExists -teamName $unitName
        if ($teamExists) {
            Write-Log "$unitName already exists"
            # Get the Microsoft 365 Group associated with the Team
            $group = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$unitName' and resourceProvisioningOptions/Any(x:x eq 'Team')"
            $groupId = $group.value[0].id
        
            $groupDetails = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$groupId"
            $currentProxyAddresses = $groupDetails.proxyAddresses | Where-Object { $_ -ne $null -and $_ -ne "" }  # Filter out null or empty values
            Write-Output "Current Proxy Addresses: $currentProxyAddresses"
        
            $alias = "smtp:" + (New-TeamAlias -unitName $unitName)
            Write-Log "Generated alias for $unitName : $alias"
        
            # Check if the alias exists
            if ($currentProxyAddresses -contains $alias) {
                Write-Log "Alias $alias already exists for the Team $unitName."
            } else {
                # Check if alias is already used by another group
                $aliasInUse = $false
                try {
                    $aliasCheckUri = "https://graph.microsoft.com/v1.0/groups?\$filter=proxyAddresses/any(x:x eq '$alias')"
                    $aliasCheckResponse = Invoke-MgGraphRequest -Method GET -Uri $aliasCheckUri
                    if ($aliasCheckResponse.value.Count -gt 0) {
                        $aliasInUse = $true
                    }
                } catch {
                    Write-Log "Could not check if alias $alias is in use. Error: $_"
                }
                if ($aliasInUse) {
                    Write-Log "Alias $alias is already used by another group. Skipping adding this alias."
                } else {
                    try {
                        Set-UnifiedGroup -Identity $unitName -EmailAddresses @{Add=$alias}
                        Write-Output "Alias $alias added successfully to the Team '$unitName'."
                    } catch {
                        Write-Error "Failed to add alias $alias to the Team '$unitName'. Error: $_"
                    }
                }
            }
            $unitNumber = $unitName.Substring(3, 3)  # gets first 6 characters of string (co-XXX)
            $unitCommanderCAPID = GetCommander -allUsers $allUsers -unit $unitNumber
            $unitCommander = $allUsers | Where-Object { $_.officeLocation -eq $unitCommanderCAPID.CAPID } 
            # check if team owner is the same as the commander
            $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$unitName' and resourceProvisioningOptions/Any(x:x eq 'Team')"
            # Make the API request
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri

            $teamId = $response.value[0].id
            Write-Log "Team ID: $teamId"
            $uri = "https://graph.microsoft.com/v1.0/groups/$teamId/owners"
            $teamOwners = Invoke-MgGraphRequest -Method GET -Uri $uri
            $teamOwnerIds = $teamOwners.value | ForEach-Object { $_.id }
            Write-Log "Team Owner IDs: $teamOwnerIds"
            # Check to see if the unit commander is an owner of the team
            $isOwner = $teamOwnerIds | Where-Object { $_ -eq $unitCommander.id }
            if ($isOwner) {
                Write-Log "Team Owner is the same as the Commander"
            } else {
                Write-Log "Team Owner is NOT the same as the Commander"
                # Change the team owner to the commander
                try {
                    # Check if commander is a guest
                    $commanderDetails = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$($unitCommander.id)"
                    if ($commanderDetails.userType -eq "Guest") {
                        Write-Log "Cannot add commander $($unitCommander.displayName) as team owner: user is a guest."
                    } else {
                        # Prepare the request body
                        $body = @{
                            "@odata.type" = "#microsoft.graph.aadUserConversationMember"
                            roles = @("owner")
                            "user@odata.bind" = "https://graph.microsoft.com/v1.0/users/$($unitCommander.id)"
                        } | ConvertTo-Json

                        # Add the user as an owner
                        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/members" -Body $body -ContentType "application/json"
                        Write-Log "Team owner changed to $($unitCommander.displayName)"
                    }
                } catch {
                    Write-Error "Failed to change team owner: $_"
                }
            }
            # Ensure Mike.schulte is an owner of the team so the script can run
            $isOwner = $teamOwnerIds | Where-Object { $_ -eq '53e55cd9-1413-4275-8f84-902dd4d8c0a7' }
            if ($isOwner) {
                Write-Log "Mike Schulte is an owner of the team"
            } else {
                Write-Log "Mike Schulte is NOT an owner of the team"
                # Add Mike Schulte as an owner
                try {
                # Prepare the request body
                $body = @{
                        "@odata.type" = "#microsoft.graph.aadUserConversationMember"
                        roles = @("owner")
                        "user@odata.bind" = "https://graph.microsoft.com/v1.0/users/53e55cd9-1413-4275-8f84-902dd4d8c0a7"
                    } | ConvertTo-Json
                
                # Add the user as an owner
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/members" -Body $body -ContentType "application/json"
                } catch {
                    Write-Error "Failed to add Mike Schulte as team owner: $_"
                }
            }
            # Remove Managed Identity owner logic (not supported by Teams/Graph API)
        }
        else {
            Write-Log "$unitName needs to be created"
            $unitNumber = $unitName.Substring(3, 3)      
            $unitCommander = GetCommander -allUsers $allUsers -unit $unitNumber
            Write-Log "Commander is $($unitCommander.displayName)"
            $mailName = "CO" + $unitNumber
            Write-Host "mailname: $mailName"
            $commanderId = $unitCommander.mail
            Write-Log $commanderId
            $uri = "https://graph.microsoft.com/v1.0/users('$commanderId')"
            $mikeSchulteId = '53e55cd9-1413-4275-8f84-902dd4d8c0a7'
            $mikeSchulteUri = "https://graph.microsoft.com/v1.0/users/$mikeSchulteId"
            # Always add Mike Schulte as owner, add commander if found
            $membersArr = @(
                @{
                    "@odata.type" = "#microsoft.graph.aadUserConversationMember"
                    roles = @("owner")
                    "user@odata.bind" = $mikeSchulteUri
                }
            )
            if ($commanderId) {
                $membersArr += @{
                    "@odata.type" = "#microsoft.graph.aadUserConversationMember"
                    roles = @("owner","member")
                    "user@odata.bind" = $uri
                }
            }
            try {
                $teamPayload = @{
                    "template@odata.bind" = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
                    displayName = $unitName
                    mailNickname = $mailName
                    description = "Team for $unitName"
                    visibility = "Private"
                    members = $membersArr
                }
                $response = New-MgTeam -BodyParameter $teamPayload
                if ($null -ne $response) {
                    Write-Log "Team created successfully. Team ID: $($response.id)"
                } else {
                    Write-Log "Team creation failed. No response received."
                }
            } catch {
                Write-Error "An error occurred: $_"
                Write-Error "Error Details: $($_.Exception.Message)"                
            }
        }
    }
}

function PopulateTeams {
    param (
        [array]$allUsers,
        [array]$unitList 
    )
    
    foreach ($unitName in $unitList) {
        Write-Log "unitName is: $unitName"
        # Define the endpoint to get the team by display name
        $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=startswith(displayName,'$unitName') and resourceProvisioningOptions/Any(x:x eq 'Team')"
        $team = Invoke-MgGraphRequest -Method GET -Uri $uri
        $teamId = $team.value[0].id
        $uri = "https://graph.microsoft.com/v1.0/groups/$teamId"

        if (-not [string]::IsNullOrEmpty($teamId)) {

            Write-Log "Team Found: $teamId"

            $unitNumber = $unitName.Substring(0, 6) # get the "CO-XXX"
            $unitMembers = @()
            $unitMembers = $allUsers | Where-Object { $_.companyName -eq $unitNumber } # gets everyone in EntraID in the unit
            # Write-Log $unitMembers
            $unitMemberIds = $unitMembers | Select-Object -ExpandProperty ID    
            # Normalize and remove duplicates from unitMemberIds
            $unitMemberIds = $unitMemberIds | Sort-Object -Unique

            # Get all members on the Team
            $apiEndpoint = "https://graph.microsoft.com/v1.0/teams/$teamId/members"
            $teamMembers = @()
            do {
                $response = Invoke-MgGraphRequest -Method GET -Uri $apiEndpoint
                $teamMembers += $response.value
                $apiEndpoint = $response.'@odata.nextLink'
            } while ($apiEndpoint)

            # Extract and normalize team user IDs
            $teamUserIds = $teamMembers | ForEach-Object { $_.userId } | Sort-Object -Unique

            # Debug output
            Write-Host "Unit Member IDs: $($unitMemberIds.Count)"
            Write-Host "Team User IDs: $($teamUserIds.Count)"

            # Compare $unitMemberIds and $teamUserIds
            $result = Compare-UserIds -Array1 $unitMemberIds -Array2 $teamUserIds

            # Debug comparison results
            Write-Host "Users to Add to Team:"
            $result.AddtoTeams | ForEach-Object { Write-Host $_ }

            Write-Host "Users to Remove from Team:"
            $result.RemovefromTeams | ForEach-Object { Write-Host $_ }

            # Add each user in AddtoTeams to the team
            foreach ($userId in $result.AddtoTeams) {
                if ($teamUserIds -contains $userId) {
                    Write-Host "Skipping user $userId as they are already in the team."
                    continue
                }

                # Check if user exists in Entra ID before proceeding
                $userExists = $true
                try {
                    $userDetails = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$userId"
                } catch {
                    if ($_.Exception.Response -and ($_.Exception.Response.StatusCode.value__ -eq 404)) {
                        Write-Log "User $userId does not exist in Entra ID. Skipping."
                        $userExists = $false
                    } else {
                        Write-Log "Could not retrieve user details for $userId. Skipping. Error: $_"
                        $userExists = $false
                    }
                }
                if (-not $userExists) {
                    continue
                }

                # Additional check: skip if account is not enabled
                if ($userDetails.accountEnabled -eq $false) {
                    Write-Log "Skipping user $userId : account is disabled."
                    continue
                }

                # Log userType for debugging
                Write-Log "User $userId userType: $($userDetails.userType)"

                # Skip license check, assume all users are guests
                Write-Log "Adding user $userId as a guest member (license check skipped)."

                try {
                    Write-Log "Adding User: $userId to the Team"
                    # Define the API endpoint for adding a member to the Team
                    $apiEndpoint = "https://graph.microsoft.com/v1.0/teams/$teamId/members"
                    $body = @{
                        "@odata.type" = "#microsoft.graph.aadUserConversationMember"
                        roles         = @()  # Always member for guests
                        "user@odata.bind" = "https://graph.microsoft.com/v1.0/users/$userId"
                    } | ConvertTo-Json -Depth 10

                    Invoke-MgGraphRequest -Method POST -Uri $apiEndpoint -Body $body -ContentType "application/json"
                    Write-Log "User $userId added successfully to the Team."
                } catch {
                    if ($_.Exception.Response -and ($_.Exception.Response.StatusCode.value__ -eq 403)) {
                        Write-Error "Forbidden: You do not have permission to add user $userId to the Team. This may be due to licensing, guest status, or lack of permissions."
                    } elseif ($_.Exception.Response -and ($_.Exception.Response.StatusCode.value__ -eq 404)) {
                        Write-Log "User $userId not found when adding to Team. Skipping."
                    } else {
                        Write-Error "Failed to add user $userId to the Team. Error: $_"
                        Write-Error "Error Details: $($_.Exception.Message)"
                    }
                }
            }
            # Delete users in RemovefromTeams
            # foreach ($userId in $result.RemovefromTeams) {
            #     try {
            #         # Define the API endpoint to get team members
            #         $apiEndpoint = "https://graph.microsoft.com/v1.0/teams/$teamId/members"

            #         # Fetch all members of the team
            #         $teamMembers = @()
            #         do {
            #             $response = Invoke-MgGraphRequest -Method GET -Uri $apiEndpoint
            #             $teamMembers += $response.value
            #             $apiEndpoint = $response.'@odata.nextLink'
            #         } while ($apiEndpoint)

            #         # Filter the members to find the one with the specific userId
            #         $membership = $teamMembers | Where-Object { $_.userId -eq $userId }               
            #         if ($null -ne $membership) {
            #             # Remove the member from the team
            #             $apiEndpoint = "https://graph.microsoft.com/v1.0/teams/$teamId/members/$($membership.id)"
            #             try {
            #                 # Make the DELETE request to remove the member
            #                 Invoke-MgGraphRequest -Method DELETE -Uri $apiEndpoint
            #                 Write-Log "User with Membership ID $($membership.displayName) removed successfully from Team $teamId."
            #             } catch {
            #                 Write-Error "Failed to remove user with Membership ID $($membership.displayName) from Team $teamId. Error: $_"
            #             }
            #         }
            #     } catch {
            #         Write-Error "Failed to remove user with Membership ID $($membership.displayName) from Team $teamId. Error: $_"
            #     }
            # }
        }
    }
}

# Main Program starts here...
# ---------------------------

Write-Log "We are in the Teams script"
$unitList = GetUnits  # gets all units
$allGroups = GetAllGroups
$allUsers = GetAllUsers
CheckTeams -unitList $unitList -allGroups $allGroups -allUsers $allUsers # gets all teams and creates teams if needed
PopulateTeams -allUsers $allUsers -unitList $unitList
Write-Log "Teams script completed successfully."