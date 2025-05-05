# Input bindings are passed in via param block.
param($Timer)

# Include shared Functions
. "$PSScriptRoot\..\shared\shared.ps1"

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
            Write-Output "User ID: $userId, Display Name: $displayName"
        } catch {
            Write-Error "Failed to retrieve display name for User ID: $userId. Error: $_"
        }
    }
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
            Write-Output "$unitName already exists"
            Write-Output "Checking if the alias exists and then creating it if not"
            $alias = New-TeamAlias -unitName $unitName
            Write-Output "Generated alias for $unitName : $alias"
            # Check if the alias exists
            if ($group.EmailAddresses -contains $alias) {
                Write-Output "Alias $alias already exists for the Team $teamName."
            } else {
            # Add the alias to the group
            $body = @{
                proxyAddresses = @("smtp:$alias")
            } | ConvertTo-Json -Depth 10
            try {
                $groupId = $group.id
                Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/groups/$groupId" -Body $body -ContentType "application/json"
                Write-Output "Alias $alias added successfully to the Team '$teamName'."
            } catch {
                Write-Error "Failed to add alias $alias to the Team '$teamName'. Error: $_"
            }            }    
            $unitNumber = $unitName.Substring(3, 3)  # gets first 6 characters of string (co-XXX)
            $unitCommanderCAPID = GetCommander -allUsers $allUsers -unit $unitNumber
            $unitCommander = $allUsers | Where-Object { $_.officeLocation -eq $unitCommanderCAPID.CAPID } 
            # check if team owner is the same as the commander
            $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$unitName' and resourceProvisioningOptions/Any(x:x eq 'Team')"
            # Make the API request
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri

            $teamId = $response.value[0].id
            Write-Output "Team ID: $teamId"
            $uri = "https://graph.microsoft.com/v1.0/groups/$teamId/owners"
            $teamOwners = Invoke-MgGraphRequest -Method GET -Uri $uri
            $teamOwnerIds = $teamOwners.value | ForEach-Object { $_.id }
            Write-Output "Team Owner IDs: $teamOwnerIds"
            # Check to see if the unit commander is an owner of the team
            $isOwner = $teamOwnerIds | Where-Object { $_ -eq $unitCommander.id }
            if ($isOwner) {
                Write-Output "Team Owner is the same as the Commander"
            } else {
                Write-Output "Team Owner is NOT the same as the Commander"
                # Change the team owner to the commander
                try {
                # Prepare the request body
                $body = @{
                        "@odata.type" = "#microsoft.graph.aadUserConversationMember"
                        roles = @("owner")
                        "user@odata.bind" = "https://graph.microsoft.com/v1.0/users/$($unitCommander.id)"
                    } | ConvertTo-Json
                
                # Add the user as an owner
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/members" -Body $body -ContentType "application/json"
                # Remove the old owner              
                    Write-Output "Team owner changed to $($teamOwner.displayName)"
                } catch {
                    Write-Error "Failed to change team owner: $_"
                }
            }
            # Ensure Mike.schulte is an owner of the team so the script can run
            $isOwner = $teamOwnerIds | Where-Object { $_ -eq '53e55cd9-1413-4275-8f84-902dd4d8c0a7' }
            if ($isOwner) {
                Write-Output "Mike Schulte is an owner of the team"
            } else {
                Write-Output "Mike Schulte is NOT an owner of the team"
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
        } else {
            Write-Output "$unitName needs to be created"
            $unitNumber = $unitName.Substring(3, 3)  # gets first 6 characters of string (co-XXX)
            Write-Host $unitNumber        
            $unitCommander = GetCommander -allUsers $allUsers -unit $unitNumber # gets first instance of commander
            Write-Host $unitCommander.displayName
            $mailName = "CO" + $unitNumber # makes mail for Team 'coxxx@cowg.cap.gov'
            Write-Host "mailname: $mailName"
            $commanderId = $unitCommander.mail # commander's email to make owner of team
            Write-Output $commanderId
            $uri = "https://graph.microsoft.com/v1.0/users('$commanderId')"
            if ($commanderId) {
                try {
                    # Create the team using Microsoft Graph API
                    $teamPayload = @{
                        "template@odata.bind" = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
                        displayName = $unitName
                        mailNickname = $mailName
                        description = "Team for $unitName"
                        visibility = "Private"
                        members = @(
                            @{
                                "@odata.type" = "#microsoft.graph.aadUserConversationMember"
                                roles = @(
                                "owner","member"
                            )
                            "user@odata.bind" = $uri
                        }
                    )                    
                }
                $response = New-MgTeam -BodyParameter $teamPayload
                    # Check if the response is not null
                if ($null -ne $response) {
                    Write-Output "Team created successfully. Team ID: $($response.id)"
                } else {
                    Write-Output "Team creation failed. No response received."
                }
                    # Verify the team creation by querying the team
                } catch {
                    Write-Error "An error occurred: $_"
                    Write-Error "Error Details: $($_.Exception.Message)"                
                }
            } else {
                Write-Output "No Commander found for $unitNumber"
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
        Write-Output "unitName is: $unitName"
        # Define the endpoint to get the team by display name
        $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=startswith(displayName,'$unitName') and resourceProvisioningOptions/Any(x:x eq 'Team')"
        $team = Invoke-MgGraphRequest -Method GET -Uri $uri
        $teamId = $team.value[0].id
        $uri = "https://graph.microsoft.com/v1.0/groups/$teamId"

        if (-not [string]::IsNullOrEmpty($teamId)) {

            Write-Output "Team Found: $teamId"

            $unitNumber = $unitName.Substring(0, 6) # get the "CO-XXX"
            $unitMembers = @()
            $unitMembers = $allUsers | Where-Object { $_.companyName -eq $unitNumber } # gets everyone in EntraID in the unit
            # Write-Output $unitMembers
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
            
                try {
                    Write-Output "Adding User: $userId to the Team"
            
                    # Define the API endpoint for adding a member to the Team
                    $apiEndpoint = "https://graph.microsoft.com/v1.0/teams/$teamId/members"
            
                    # Prepare the request body
                    $body = @{
                        "@odata.type" = "#microsoft.graph.aadUserConversationMember"
                        roles         = @()  # Empty array for a member (use @("owner") for an owner)
                        "user@odata.bind" = "https://graph.microsoft.com/v1.0/users/$userId"
                    } | ConvertTo-Json -Depth 10
            
                    # Make the API request to add the user to the Team
                    Invoke-MgGraphRequest -Method POST -Uri $apiEndpoint -Body $body -ContentType "application/json"
            
                    Write-Output "User $userId added successfully to the Team."
                } catch {
                    Write-Error "Failed to add user $userId to the Team. Error: $_"
                    Write-Error "Error Details: $($_.Exception.Message)"
                }
            }

            # Delete users in RemovefromTeams
            foreach ($userId in $result.RemovefromTeams) {
                try {
                    # Define the API endpoint to get team members
                    $apiEndpoint = "https://graph.microsoft.com/v1.0/teams/$teamId/members"

                    # Fetch all members of the team
                    $teamMembers = @()
                    do {
                        $response = Invoke-MgGraphRequest -Method GET -Uri $apiEndpoint
                        $teamMembers += $response.value
                        $apiEndpoint = $response.'@odata.nextLink'
                    } while ($apiEndpoint)

                    # Filter the members to find the one with the specific userId
                    $membership = $teamMembers | Where-Object { $_.userId -eq $userId }               
                    if ($null -ne $membership) {
                        # Remove the member from the team
                        $apiEndpoint = "https://graph.microsoft.com/v1.0/teams/$teamId/members/$($membership.id)"
                        try {
                            # Make the DELETE request to remove the member
                            Invoke-MgGraphRequest -Method DELETE -Uri $apiEndpoint
                            Write-Output "User with Membership ID $($membership.displayName) removed successfully from Team $teamId."
                        } catch {
                            Write-Error "Failed to remove user with Membership ID $($membership.displayName) from Team $teamId. Error: $_"
                        }
                    }
                } catch {
                    Write-Error "Failed to remove user with Membership ID $($membership.displayName) from Team $teamId. Error: $_"
                }
            }
        }
    }
}


function UpdateAnnouncements {
    param (
        [array]$unitList 
    )
    foreach ($unitName in $unitList) {
        Write-Output "unitName is: $unitName"
        # Define the endpoint to get the team by display name
        $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=startswith(displayName,'$unitName') and resourceProvisioningOptions/Any(x:x eq 'Team')"
        $team = Invoke-MgGraphRequest -Method GET -Uri $uri
        $teamValue = $team.value[0]
        $teamMail = $teamValue.mail
        $uri = "https://graph.microsoft.com/v1.0/groups/$teamMail"

        $distributionList = "announcements@cowg.cap.gov"
        Add-DistributionGroupMember -Identity $distributionList -Member $teamMail
        Write-Output "Team: $teamMail added to announcements"
    }
}

# Main Program starts here...
# ---------------------------
Write-Log "Starting UpdateTeams script execution----------------------------------------------------"
Write-Output "We are in main"
$unitList = GetUnits  # gets all units
$allGroups = GetAllGroups
$allUsers = GetAllUsers
CheckTeams -unitList $unitList -allGroups $allGroups -allUsers $allUsers # gets all teams and creates teams if needed
PopulateTeams -allUsers $allUsers -unitList $unitList
# UpdateAnnouncements -unitList $unitList
Write-Log "UpdateTeams script execution completed----------------------------------------------------"