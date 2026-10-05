param($Timer)

# Connect to Microsoft Graph
$MSGraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token
Connect-MgGraph -AccessToken $MSGraphAccessToken -NoWelcome

# Include shared Functions
. "$PSScriptRoot\..\shared\shared.ps1"


# Resolve script root robustly (Azure Functions may not populate MyInvocation.MyCommand.Path)
$ScriptRoot = $PSScriptRoot
if (-not $ScriptRoot -or $ScriptRoot -eq '') { try { $ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path } catch { $ScriptRoot = (Get-Location).Path } }
Write-Log "ScriptRoot resolved to: $ScriptRoot"

# Timer-triggered wrapper for syncing one or more distribution groups to Teams
# Expectations / assumptions:
# - A CSV file named 'group_to_team.csv' may be placed next to this script with columns: GroupId,TeamDisplayName
# - If the CSV is missing, you can set the environment variable GROUP_TEAM_PAIRS with semicolon-separated pairs
#   e.g. GROUP_TEAM_PAIRS="<groupId1>:Team Name 1;<groupId2>:Team Name 2"
# - By default the function does a dry-run and writes per-team CSVs into ../output
# - To actually add members set env var EXECUTE=true. To skip prompts set FORCE=true.
# - To create missing Teams set CREATE_MISSING_TEAMS=true. Optionally limit creation with CREATE_MISSING_TEAM_UNITS=200,141.

function Sync-GroupToTeam {
    param(
        [string]$GroupId,
        [string]$TeamDisplayName,
        [bool]$Execute = $true,
        [bool]$Force = $true
    )

    Write-Log "Sync start: GroupId=$GroupId -> Team='$TeamDisplayName' (Execute=$Execute, Force=$Force)"
    # Using the existing Graph connection available in the runspace (no Connect-GraphIfNeeded helper required)

    # Fetch group members (users only)
    $members = @()
    $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/members/microsoft.graph.user?`$select=id,displayName,mail,userPrincipalName"
    $headers = @{ 'ConsistencyLevel' = 'eventual' }
    try {
        do {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $headers
            if ($resp -and $resp.value) { $members += $resp.value }
            $uri = $resp.'@odata.nextLink'
        } while ($uri)
    } catch {
        Write-Log ("Failed to retrieve group members for {0}: {1}" -f $GroupId, $_)
        return
    }

    if ($members.Count -eq 0) { Write-Log "No user members found in group $GroupId"; return }

    # Find the Team-enabled Microsoft 365 group by display name. Distribution groups
    # can share the same display name as the Team, so filter for Teams explicitly.
    $team = $null
    try {
        $escapedTeamDisplayName = $TeamDisplayName.Replace("'", "''")
        $filter = "displayName eq '$escapedTeamDisplayName'"
        $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=$([uri]::EscapeDataString($filter))&`$select=id,displayName,mail,resourceProvisioningOptions"
        $teams = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $headers
        $team = @($teams.value | Where-Object { Test-TeamResourceProvisioningOption -Group $_ }) | Select-Object -First 1
    } catch {
        Write-Log ("Team lookup failed for {0}: {1}" -f $TeamDisplayName, $_)
        $team = $null
    }
    if (-not $team) { Write-Log "Team with display name '$TeamDisplayName' not found."; return }
    $teamId = $team.id
    Write-Log "Team found: $($team.displayName) ($teamId)"

    # Fetch current team members (group members)
    $currentMembers = @()
    $uri = "https://graph.microsoft.com/v1.0/groups/$teamId/members/microsoft.graph.user?`$select=id,displayName,mail,userPrincipalName"
    do {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $headers
        if ($resp -and $resp.value) { $currentMembers += $resp.value }
        $uri = $resp.'@odata.nextLink'
    } while ($uri)

    # Determine which members to add (by user id)
    $currentIds = $currentMembers | ForEach-Object { $_.id }
    $toAdd = $members | Where-Object { $currentIds -notcontains $_.id }

    # Log preview of changes
    if ($Execute) {
        Write-OperationLog -Operation "Adding members to Team" -Details "$($toAdd.count) members for '$TeamDisplayName'"
    } else {
        Write-OperationLog -Operation "Would add members to Team" -Details "$($toAdd.count) members for '$TeamDisplayName'"
    }
    
    if ($toAdd.Count -gt 0) {
        foreach ($u in $toAdd[0..([Math]::Min(9,$toAdd.Count-1))]) {
            $detail = "{0} <{1}> ({2})" -f $u.displayName, $u.mail, $u.userPrincipalName
            if ($Execute) {
                Write-OperationLog -Operation "Candidate for addition" -Details $detail
            } else {
                Write-OperationLog -Operation "Would add candidate" -Details $detail
            }
        }
        if ($toAdd.Count -gt 10) { 
            Write-Log "...and $($toAdd.Count - 10) more candidates not shown" 
        }
    }

    if ($Execute -and $toAdd.Count -gt 0) {
        Write-Log "Executing: adding $($toAdd.count) users to Team '$TeamDisplayName'"
        foreach ($u in $toAdd) {
            $display = "$($u.displayName) <$($u.mail) | $($u.userPrincipalName)>"
            $doIt = $false
            if ($Force) { $doIt = $true } elseif (-not $Host.UI.RawUI.KeyAvailable) { # Non-interactive host
                Write-Log "Non-interactive host detected; skipping interactive confirmation for $display"
                $doIt = $false
            } else { $confirm = Read-Host "Add $display to Team? (Y/N)"; if ($confirm -match '^[Yy]') { $doIt = $true } }
            if (-not $doIt) { Write-Log "Skipped: $display"; continue }

            # Add member via /groups/{team-id}/members/$ref with @odata.id
            $memberRef = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/users/$($u.id)" } | ConvertTo-Json
            try {
                $postUri = "https://graph.microsoft.com/v1.0/groups/$teamId/members/`$ref"
                Invoke-MgGraphRequest -Method POST -Uri $postUri -Body $memberRef -ContentType "application/json" -Headers $headers
                Write-Log ("Added: {0}" -f $display)
            } catch {
                Write-Log ("Failed to add {0}: {1}" -f $display, $_)
            }
        }
    } elseif ($Execute) {
        Write-Log "Team '$TeamDisplayName' already has all desired members. No additions needed."
    } else {
        Write-Log "Dry-run complete for Team '$TeamDisplayName'. Re-run with EXECUTE=true to actually add members."
    }
}

function Test-TeamResourceProvisioningOption {
    param (
        [object]$Group
    )

    if (-not $Group -or -not $Group.resourceProvisioningOptions) {
        return $false
    }

    return @($Group.resourceProvisioningOptions) -contains "Team"
}

function Get-AllTeamGroups {
    $headers = @{ 'ConsistencyLevel' = 'eventual' }
    $teams = @()
    $uri = "https://graph.microsoft.com/v1.0/groups?`$select=id,displayName,mail,resourceProvisioningOptions&`$top=999"

    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $headers
        if ($response -and $response.value) {
            $teams += @($response.value | Where-Object { Test-TeamResourceProvisioningOption -Group $_ })
        }
        $uri = $response.'@odata.nextLink'
    } while ($uri)

    return @($teams | Sort-Object displayName)
}

function Get-OrganizationTeamPrefix {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Organization
    )

    return "$($env:WING_DESIGNATOR)-$($Organization.Unit)"
}

function Find-TeamGroupsForOrganization {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Organization,

        [Parameter(Mandatory = $true)]
        [array]$TeamGroups
    )

    $prefix = Get-OrganizationTeamPrefix -Organization $Organization
    $escapedPrefix = [regex]::Escape($prefix)

    return @(
        $TeamGroups |
            Where-Object { $_.displayName -match "^$escapedPrefix($|[^0-9])" } |
            Sort-Object displayName
    )
}

function Get-BooleanEnvironmentFlag {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [bool]$DefaultValue = $false
    )

    $value = [Environment]::GetEnvironmentVariable($Name)
    if ([string]::IsNullOrWhiteSpace($value)) {
        return $DefaultValue
    }

    return $value -match '^(true|yes|1)$'
}

function Get-TeamDisplayNameForOrganization {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Organization
    )

    $prefix = Get-OrganizationTeamPrefix -Organization $Organization
    if ([string]::IsNullOrWhiteSpace($Organization.Name)) {
        return $prefix
    }

    return "$prefix $($Organization.Name)"
}

function Get-TeamCreationUnitFilter {
    $value = [Environment]::GetEnvironmentVariable("CREATE_MISSING_TEAM_UNITS")
    if ([string]::IsNullOrWhiteSpace($value)) {
        return @()
    }

    return @(
        $value -split '[,; ]+' |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object {
                if ($_ -match '(\d+)') {
                    $matches[1].PadLeft(3, '0')
                }
            } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
    )
}

function Test-TeamCreationUnitIncluded {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Organization,

        [string[]]$UnitFilter = @()
    )

    if (-not $UnitFilter -or $UnitFilter.Count -eq 0) {
        return $true
    }

    return $UnitFilter -contains $Organization.Unit
}

function Get-CapwatchCommandersPath {
    param (
        [string]$CapwatchDataDir = (Join-Path $env:HOME "data/CAPWatch")
    )

    return Join-Path -Path $CapwatchDataDir -ChildPath "Commanders.txt"
}

function Import-CapwatchCommanderRows {
    param (
        [string]$Path,
        [string]$CapwatchDataDir = (Join-Path $env:HOME "data/CAPWatch")
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        $Path = Get-CapwatchCommandersPath -CapwatchDataDir $CapwatchDataDir
    }

    if (-not (Test-Path -Path $Path)) {
        throw "CAPWATCH Commanders.txt was not found at $Path"
    }

    return @(Import-Csv -Path $Path -ErrorAction Stop)
}

function Get-CapwatchCadetDutyPositionsPath {
    param (
        [string]$CapwatchDataDir = (Join-Path $env:HOME "data/CAPWatch")
    )

    return Join-Path -Path $CapwatchDataDir -ChildPath "CadetDutyPositions.txt"
}

function Import-CapwatchCadetDutyPositionRows {
    param (
        [string]$Path,
        [string]$CapwatchDataDir = (Join-Path $env:HOME "data/CAPWatch")
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        $Path = Get-CapwatchCadetDutyPositionsPath -CapwatchDataDir $CapwatchDataDir
    }

    if (-not (Test-Path -Path $Path)) {
        throw "CAPWATCH CadetDutyPositions.txt was not found at $Path"
    }

    return @(Import-Csv -Path $Path -ErrorAction Stop)
}

function Get-CapwatchDutyPositionsPath {
    param (
        [string]$CapwatchDataDir = (Join-Path $env:HOME "data/CAPWatch")
    )

    return Join-Path -Path $CapwatchDataDir -ChildPath "DutyPosition.txt"
}

function Import-CapwatchDutyPositionRows {
    param (
        [string]$Path,
        [string]$CapwatchDataDir = (Join-Path $env:HOME "data/CAPWatch")
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        $Path = Get-CapwatchDutyPositionsPath -CapwatchDataDir $CapwatchDataDir
    }

    if (-not (Test-Path -Path $Path)) {
        throw "CAPWATCH DutyPosition.txt was not found at $Path"
    }

    return @(Import-Csv -Path $Path -ErrorAction Stop)
}

function Get-CacTeamDisplayName {
    $teamName = [Environment]::GetEnvironmentVariable("CAC_TEAM_DISPLAY_NAME")
    if ([string]::IsNullOrWhiteSpace($teamName)) {
        return "CAC"
    }

    return $teamName.Trim()
}

function Get-CacCadetMemberCapIds {
    param (
        [array]$CadetDutyPositions = @()
    )

    return @(
        $CadetDutyPositions |
            Where-Object {
                "$($_.FunctArea)".Trim() -match '(?i)^CAC Appointment$' -or
                "$($_.Duty)".Trim() -match '(?i)\bCAC\b'
            } |
            Select-Object -ExpandProperty CAPID |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { "$_".Trim() } |
            Sort-Object -Unique
    )
}

function Get-CacSeniorOwnerCapIds {
    param (
        [array]$DutyPositions = @()
    )

    return @(
        $DutyPositions |
            Where-Object {
                "$($_.Duty)".Trim() -match '(?i)^Cadet Advisory Council Senior Advisor$'
            } |
            Select-Object -ExpandProperty CAPID |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { "$_".Trim() } |
            Sort-Object -Unique
    )
}

function Resolve-GraphUsersByCapIds {
    param (
        [array]$CapIds = @(),
        [string]$Purpose = "CAPWATCH assignment"
    )

    $users = @()
    foreach ($capId in @($CapIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)) {
        $user = Get-GraphUserByCapid -Capid $capId
        if (-not $user -or [string]::IsNullOrWhiteSpace($user.id)) {
            Write-Log "$Purpose CAPID $capId was not found in Microsoft Graph."
            continue
        }

        $users += $user
    }

    return @($users | Sort-Object id -Unique)
}

function Get-TeamUserRelationship {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TeamId,

        [Parameter(Mandatory = $true)]
        [ValidateSet("members", "owners")]
        [string]$Relationship
    )

    $users = @()
    $headers = @{ 'ConsistencyLevel' = 'eventual' }
    $uri = "https://graph.microsoft.com/v1.0/groups/$TeamId/$Relationship/microsoft.graph.user?`$select=id,displayName,mail,userPrincipalName,employeeId"

    do {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $headers
        if ($resp -and $resp.value) {
            $users += $resp.value
        }
        $uri = $resp.'@odata.nextLink'
    } while ($uri)

    return @($users)
}

function Add-GraphUserToTeamGroupRelationship {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TeamId,

        [Parameter(Mandatory = $true)]
        [object]$User,

        [Parameter(Mandatory = $true)]
        [ValidateSet("members", "owners")]
        [string]$Relationship,

        [bool]$Execute = $false
    )

    $operation = if ($Relationship -eq "owners") { "CAC Team owner" } else { "CAC Team member" }
    $details = "$($User.displayName) (CAPID $($User.employeeId))"

    if ($Execute) {
        Write-OperationLog -Operation "Adding $operation" -Details $details
        try {
            $memberRef = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/users/$($User.id)" } | ConvertTo-Json
            Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups/$TeamId/$Relationship/`$ref" -Body $memberRef -ContentType "application/json" -Headers @{ 'ConsistencyLevel' = 'eventual' }
            Write-Log "Added $($User.displayName) as $operation."
        } catch {
            Write-Log "Failed to add $($User.displayName) as $operation`: $_"
        }
    } else {
        Write-OperationLog -Operation "Would add $operation" -Details $details
    }
}

function Remove-GraphUserFromTeamGroupRelationship {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TeamId,

        [Parameter(Mandatory = $true)]
        [object]$User,

        [Parameter(Mandatory = $true)]
        [ValidateSet("members", "owners")]
        [string]$Relationship,

        [bool]$Execute = $false
    )

    $operation = if ($Relationship -eq "owners") { "CAC Team owner" } else { "CAC Team member" }
    $details = "$($User.displayName) <$($User.mail)>"

    if ($Execute) {
        Write-OperationLog -Operation "Removing $operation" -Details $details
        try {
            Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/groups/$TeamId/$Relationship/$($User.id)/`$ref" -Headers @{ 'ConsistencyLevel' = 'eventual' }
            Write-Log "Removed $($User.displayName) as $operation."
        } catch {
            Write-Log "Failed to remove $($User.displayName) as $operation`: $_"
        }
    } else {
        Write-OperationLog -Operation "Would remove $operation" -Details $details
    }
}

function Sync-CacTeamFromCapwatch {
    param (
        [array]$CadetDutyPositions = @(),
        [array]$DutyPositions = @(),
        [array]$TeamGroups = @(),
        [bool]$Execute = $false
    )

    $teamName = Get-CacTeamDisplayName
    $team = Find-TeamGroupByDisplayName -TeamDisplayName $teamName -TeamGroups $TeamGroups
    if (-not $team) {
        Write-Log "CAC Team '$teamName' was not found. Skipping CAC duty position sync."
        return
    }

    $memberCapIds = @(Get-CacCadetMemberCapIds -CadetDutyPositions $CadetDutyPositions)
    $ownerCapIds = @(Get-CacSeniorOwnerCapIds -DutyPositions $DutyPositions)

    Write-Log "Discovered $($memberCapIds.Count) CAC cadet CAPID(s) from CadetDutyPositions.txt and $($ownerCapIds.Count) senior advisor CAPID(s) from DutyPosition.txt."

    $desiredCadetMembers = @(Resolve-GraphUsersByCapIds -CapIds $memberCapIds -Purpose "CAC cadet")
    $desiredOwners = @(Resolve-GraphUsersByCapIds -CapIds $ownerCapIds -Purpose "CAC senior advisor")
    $desiredTeamUsers = @($desiredCadetMembers + $desiredOwners | Sort-Object id -Unique)

    $currentMembers = @(Get-TeamUserRelationship -TeamId $team.id -Relationship "members")
    $currentOwners = @(Get-TeamUserRelationship -TeamId $team.id -Relationship "owners")
    $currentMemberIds = @($currentMembers | ForEach-Object { $_.id })
    $currentOwnerIds = @($currentOwners | ForEach-Object { $_.id })
    $desiredMemberIds = @($desiredTeamUsers | ForEach-Object { $_.id })
    $desiredOwnerIds = @($desiredOwners | ForEach-Object { $_.id })

    foreach ($user in $desiredTeamUsers) {
        if ($currentMemberIds -contains $user.id) {
            Write-Log "CAC user $($user.displayName) is already a member of Team '$($team.displayName)'."
            continue
        }

        Add-GraphUserToTeamGroupRelationship -TeamId $team.id -User $user -Relationship "members" -Execute:$Execute
        $currentMemberIds += $user.id
    }

    foreach ($user in $desiredOwners) {
        if ($currentOwnerIds -contains $user.id) {
            Write-Log "CAC senior advisor $($user.displayName) is already an owner of Team '$($team.displayName)'."
            continue
        }

        Add-GraphUserToTeamGroupRelationship -TeamId $team.id -User $user -Relationship "owners" -Execute:$Execute
        $currentOwnerIds += $user.id
    }

    if ($desiredOwnerIds.Count -eq 0) {
        Write-Log "No CAC senior advisor owners resolved from DutyPosition.txt. Skipping CAC owner removals to avoid leaving Team '$($team.displayName)' without a CAPWATCH owner."
        return
    }

    foreach ($user in $currentOwners) {
        if ($desiredOwnerIds -contains $user.id) {
            continue
        }

        Remove-GraphUserFromTeamGroupRelationship -TeamId $team.id -User $user -Relationship "owners" -Execute:$Execute
    }

    foreach ($user in $currentMembers) {
        if ($desiredMemberIds -contains $user.id) {
            continue
        }

        Remove-GraphUserFromTeamGroupRelationship -TeamId $team.id -User $user -Relationship "members" -Execute:$Execute
    }
}

function Find-CapwatchCommanderForOrganization {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Organization,

        [array]$Commanders = @(),

        [string]$WingDesignator = $env:WING_DESIGNATOR
    )

    if ([string]::IsNullOrWhiteSpace($WingDesignator)) {
        $WingDesignator = "CO"
    }

    $matches = @(
        $Commanders |
            Where-Object {
                "$($_.Wing)".Trim().ToUpperInvariant() -eq $WingDesignator.Trim().ToUpperInvariant() -and
                "$($_.Unit)".Trim().PadLeft(3, '0') -eq "$($Organization.Unit)".Trim().PadLeft(3, '0') -and
                -not [string]::IsNullOrWhiteSpace($_.CAPID)
            } |
            Sort-Object DateAsg -Descending
    )

    return $matches | Select-Object -First 1
}

function Resolve-CapwatchTeamOwner {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Organization,

        [array]$Commanders = @()
    )

    $teamDisplayName = Get-TeamDisplayNameForOrganization -Organization $Organization
    $commander = Find-CapwatchCommanderForOrganization -Organization $Organization -Commanders $Commanders
    if (-not $commander) {
        Write-Log "No CAPWATCH commander found for '$teamDisplayName'. Cannot create Team because Graph requires an owner."
        return $null
    }

    $commanderName = "$($commander.Rank) $($commander.NameFirst) $($commander.NameLast)".Trim()
    $ownerUser = Get-GraphUserByCapid -Capid $commander.CAPID
    if (-not $ownerUser -or [string]::IsNullOrWhiteSpace($ownerUser.id)) {
        Write-Log "CAPWATCH commander $commanderName (CAPID $($commander.CAPID)) for '$teamDisplayName' was not found in Microsoft Graph. Cannot create Team because Graph requires an owner."
        return $null
    }

    Write-Log "Resolved Team owner for '$teamDisplayName': $($ownerUser.displayName) (CAPID $($commander.CAPID), UserId $($ownerUser.id))."
    return [PSCustomObject]@{
        Commander = $commander
        User      = $ownerUser
    }
}

function Wait-TeamGroupByDisplayName {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TeamDisplayName,

        [int]$TimeoutSeconds = 180,

        [int]$PollSeconds = 10
    )

    $headers = @{ 'ConsistencyLevel' = 'eventual' }
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)

    do {
        try {
            $filter = "displayName eq '$($TeamDisplayName.Replace("'", "''"))'"
            $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=$([uri]::EscapeDataString($filter))&`$select=id,displayName,mail,resourceProvisioningOptions"
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $headers
            $team = @($response.value | Where-Object { Test-TeamResourceProvisioningOption -Group $_ }) | Select-Object -First 1

            if ($team) {
                return $team
            }
        } catch {
            Write-Log "Team lookup for '$TeamDisplayName' is not ready yet: $_"
        }

        if ((Get-Date) -lt $deadline) {
            Start-Sleep -Seconds $PollSeconds
        }
    } while ((Get-Date) -lt $deadline)

    return $null
}

function New-CapwatchTeamForOrganization {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Organization,

        [array]$Commanders = @(),

        [bool]$Execute = $false
    )

    $teamDisplayName = Get-TeamDisplayNameForOrganization -Organization $Organization
    $description = "Civil Air Patrol Colorado Wing $($Organization.Type.ToLowerInvariant()) $($Organization.Scope.ToLowerInvariant()) synchronized from CAPWATCH."
    $owner = Resolve-CapwatchTeamOwner -Organization $Organization -Commanders $Commanders
    if (-not $owner) {
        return $null
    }

    if (-not $Execute) {
        Write-OperationLog -Operation "Would create Team" -Details "$teamDisplayName with owner $($owner.User.displayName)"
        return $null
    }

    Write-OperationLog -Operation "Creating Team" -Details "$teamDisplayName with owner $($owner.User.displayName)"

    $body = @{
        "template@odata.bind" = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
        displayName            = $teamDisplayName
        description            = $description
        visibility             = "private"
        members                = @(
            @{
                "@odata.type"     = "#microsoft.graph.aadUserConversationMember"
                roles             = @("owner")
                "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$($owner.User.id)')"
            }
        )
    } | ConvertTo-Json -Depth 10

    try {
        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/teams" -Body $body -ContentType "application/json" | Out-Null
        Write-Log "Team creation requested for '$teamDisplayName'. Waiting for provisioning to complete."
    } catch {
        Write-Log "Failed to request Team creation for '$teamDisplayName': $_"
        return $null
    }

    $team = Wait-TeamGroupByDisplayName -TeamDisplayName $teamDisplayName
    if (-not $team) {
        Write-Log "Team '$teamDisplayName' was requested but was not visible as provisioned before timeout. It may still complete asynchronously; re-run updateTeams later to sync members."
        return $null
    }

    Write-Log "Team created and ready: $($team.displayName) ($($team.id))"
    return $team
}

function Find-TeamGroupByDisplayName {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TeamDisplayName,

        [array]$TeamGroups = @()
    )

    $match = @(
        $TeamGroups |
            Where-Object { "$($_.displayName)".Trim().ToLowerInvariant() -eq $TeamDisplayName.Trim().ToLowerInvariant() }
    ) | Select-Object -First 1

    if ($match) {
        return $match
    }

    $headers = @{ 'ConsistencyLevel' = 'eventual' }
    try {
        $filter = "displayName eq '$($TeamDisplayName.Replace("'", "''"))'"
        $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=$([uri]::EscapeDataString($filter))&`$select=id,displayName,mail,resourceProvisioningOptions"
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $headers
        return @($response.value | Where-Object { Test-TeamResourceProvisioningOption -Group $_ }) | Select-Object -First 1
    } catch {
        Write-Log "Team lookup failed for manual guest Team '$TeamDisplayName': $_"
        return $null
    }
}

function Sync-ManualGuestsToTeams {
    param (
        [array]$ManualGuests = @(),
        [array]$TeamGroups = @(),
        [bool]$Execute = $false
    )

    if (-not $ManualGuests -or $ManualGuests.Count -eq 0) {
        Write-Log "No manual guest Team assignments found."
        return
    }

    $assignmentsByTeam = @{}
    foreach ($guest in $ManualGuests) {
        foreach ($teamName in (Split-ManualGuestTeams -Teams $guest.Teams)) {
            $teamKey = $teamName.ToLowerInvariant()
            if (-not $assignmentsByTeam.ContainsKey($teamKey)) {
                $assignmentsByTeam[$teamKey] = [PSCustomObject]@{
                    TeamName = $teamName
                    Guests   = @()
                }
            }
            $assignmentsByTeam[$teamKey].Guests += $guest
        }
    }

    if ($assignmentsByTeam.Count -eq 0) {
        Write-Log "ManualGuests.csv did not contain any Team assignments."
        return
    }

    $headers = @{ 'ConsistencyLevel' = 'eventual' }
    foreach ($entry in $assignmentsByTeam.Values) {
        $team = Find-TeamGroupByDisplayName -TeamDisplayName $entry.TeamName -TeamGroups $TeamGroups
        if (-not $team) {
            Write-Log "Manual guest Team '$($entry.TeamName)' was not found. Skipping $($entry.Guests.Count) guest assignment(s)."
            continue
        }

        $currentMembers = @()
        $uri = "https://graph.microsoft.com/v1.0/groups/$($team.id)/members/microsoft.graph.user?`$select=id,displayName,mail,userPrincipalName"
        try {
            do {
                $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $headers
                if ($resp -and $resp.value) { $currentMembers += $resp.value }
                $uri = $resp.'@odata.nextLink'
            } while ($uri)
        } catch {
            Write-Log "Failed to retrieve current members for manual guest Team '$($team.displayName)': $_"
            continue
        }

        $currentIds = @($currentMembers | ForEach-Object { $_.id })
        foreach ($guest in $entry.Guests) {
            $user = Get-GraphUserByCapid -Capid $guest.EmployeeId
            if (-not $user) {
                $user = Get-GraphUserByEmailAddress -EmailAddress $guest.Email
            }

            if (-not $user -or [string]::IsNullOrWhiteSpace($user.id)) {
                Write-Log "Manual guest $($guest.Email) was not found in Microsoft Graph yet. Run checkAccounts first or wait for invitation creation to complete."
                continue
            }

            if ($currentIds -contains $user.id) {
                Write-Log "Manual guest $($guest.Email) is already a member of Team '$($team.displayName)'."
                continue
            }

            $sponsorDetail = ""
            if ($guest.PSObject.Properties.Name -contains "Sponsor" -and -not [string]::IsNullOrWhiteSpace($guest.Sponsor)) {
                $sponsorDetail = " (Sponsor: $($guest.Sponsor))"
            }
            $details = "$($guest.Email) -> $($team.displayName)$sponsorDetail"
            if ($Execute) {
                Write-OperationLog -Operation "Adding manual guest to Team" -Details $details
                try {
                    $memberRef = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/users/$($user.id)" } | ConvertTo-Json
                    Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups/$($team.id)/members/`$ref" -Body $memberRef -ContentType "application/json" -Headers $headers
                    Write-Log "Added manual guest $($guest.Email) to Team '$($team.displayName)'."
                    $currentIds += $user.id
                } catch {
                    Write-Log "Failed to add manual guest $($guest.Email) to Team '$($team.displayName)': $_"
                }
            } else {
                Write-OperationLog -Operation "Would add manual guest to Team" -Details $details
            }
        }
    }
}

### Main function entry (timer-triggered)
Write-Log "updateTeams timer function invoked"

try {
    # Determine execution mode from environment
    # Default: DRY-RUN mode (safe, no changes made)
    # Set EXECUTE=true in Function App settings to actually apply changes
    $dryRun = Get-DryRunMode
    $execute = -not $dryRun
    $force = $true   # default to true so non-interactive hosts won't prompt
    if ($env:FORCE) { $force = $env:FORCE.ToLower() -eq 'true' }
    $createMissingTeams = Get-BooleanEnvironmentFlag -Name "CREATE_MISSING_TEAMS" -DefaultValue $false
    $createMissingTeamUnits = @(Get-TeamCreationUnitFilter)
    
    if ($dryRun) {
        Write-Log "🔍 [DRY-RUN MODE] Previewing changes only - no Teams will be created or modified"
    } else {
        Write-Log "⚡ [EXECUTE MODE] Will actually create/modify Teams and members"
    }
    Write-Log ("Execution flags: EXECUTE={0}, FORCE={1}, CREATE_MISSING_TEAMS={2} (set these as Function App settings to change behavior)" -f $execute, $force, $createMissingTeams)
    if ($createMissingTeamUnits.Count -gt 0) {
        Write-Log "Missing Team creation limited to CAPWATCH unit(s): $($createMissingTeamUnits -join ',')"
    }

    # Load mappings from CSV file (GroupId,TeamDisplayName)
    $mappingFile = Join-Path -Path $ScriptRoot -ChildPath 'group_to_team.csv'
    $mappings = @()
    if (Test-Path $mappingFile) {
        try { $mappings = Import-Csv -Path $mappingFile } catch { Write-Log ("Failed to read mapping file {0}: {1}" -f $mappingFile, $_) }
    } elseif ($env:GROUP_TEAM_PAIRS) {
    # parse semicolon-separated group:team pairs
    $pairs = $env:GROUP_TEAM_PAIRS -split ';' | Where-Object { $_ -match ':' }
    foreach ($p in $pairs) {
        $parts = $p -split ':'
        if ($parts.Count -ge 2) { $mappings += [PSCustomObject]@{ GroupId = $parts[0].Trim(); TeamDisplayName = ($parts[1..($parts.Count-1)] -join ':').Trim() } }
    }
    } else {
        Write-Log "No explicit group->team mappings found (expected $mappingFile or env GROUP_TEAM_PAIRS). Continuing with CAPWATCH Team discovery."
    }

    $processedTeamNames = @{}
    foreach ($map in $mappings) {
        if ($null -eq $map.GroupId -or $null -eq $map.TeamDisplayName) { Write-Log "Skipping invalid mapping entry: $map"; continue }
        Sync-GroupToTeam -GroupId $map.GroupId -TeamDisplayName $map.TeamDisplayName -Execute:$execute -Force:$force
        $processedTeamNames[$map.TeamDisplayName.ToLowerInvariant()] = $true
    }

    # Helper: resolve a distribution list (by display name) to a group id and sync it to a Team
    function Sync-DLToTeam {
    param(
        [Parameter(Mandatory=$true)][string]$DLDisplayName,
        [Parameter(Mandatory=$true)][string]$TeamDisplayName,
        [bool]$Execute = $true,
        [bool]$Force = $true
    )

    Write-Log "Resolving distribution list: $DLDisplayName"
    $headers = @{ 'ConsistencyLevel' = 'eventual' }

    # Attempt 1: lookup by mail (email address), e.g. CO-022@cowg.cap.gov
    $groupId = $null
    try {
        $filter = "mail eq '$DLDisplayName'"
        $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=$([uri]::EscapeDataString($filter))&`$select=id,displayName,mail"
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $headers
        if ($resp -and $resp.value -and $resp.value.Count -gt 0) {
            $groupId = $resp.value[0].id
            Write-Log "Found DL by mail '$DLDisplayName' -> groupId $groupId (displayName: $($resp.value[0].displayName))."
        }
    } catch {
        Write-Log "Lookup by mail failed for '$DLDisplayName': $_"
    }

    # Attempt 2: fallback to displayName match if mail lookup didn't find anything
    if (-not $groupId) {
        try {
            $filter = "displayName eq '$DLDisplayName'"
            $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=$([uri]::EscapeDataString($filter))&`$select=id,displayName,mail"
            $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $headers
            if ($resp -and $resp.value -and $resp.value.Count -gt 0) {
                $groupId = $resp.value[0].id
                Write-Log "Found DL by displayName '$DLDisplayName' -> groupId $groupId."
            }
        } catch {
            Write-Log "Lookup by displayName failed for '$DLDisplayName': $_"
        }
    }

    if (-not $groupId) {
        Write-Log "Distribution list '$DLDisplayName' not found."
        return
    }

    Write-Log "Invoking Sync-GroupToTeam for groupId $groupId (Team: $TeamDisplayName)."
    Sync-GroupToTeam -GroupId $groupId -TeamDisplayName $TeamDisplayName -Execute:$Execute -Force:$Force
}

    # Helper wrapper: Sync a flight/list (FL) to a Team. This simply calls the DL->Team helper.
    function Sync-FLToTeam {
        param(
            [Parameter(Mandatory=$true)][string]$TeamDisplayName,
            [bool]$Execute = $false,
            [bool]$Force = $false
        )

        # Derive distribution list from the beginning of the team display name (e.g. 'CO-022 ...')
        $dlPrefix = $null
        if ($TeamDisplayName -match '^(CO-\d{2,3})') { $dlPrefix = $matches[1] }
        elseif ($TeamDisplayName -match '^(CO-\d+)') { $dlPrefix = $matches[1] }
        else {
            # Try splitting on whitespace and look for CO- prefix
            $first = ($TeamDisplayName -split '\s+')[0]
            if ($first -match '^(CO-\d+)') { $dlPrefix = $matches[1] }
        }

        if (-not $dlPrefix) {
            Write-Log "Could not derive CO-XXX prefix from TeamDisplayName '$TeamDisplayName'. Skipping."
            return
        }

        $dlAddress = "$dlPrefix@cowg.cap.gov"
        Write-Log "Sync-FLToTeam derived DL '$dlAddress' from Team '$TeamDisplayName' (Execute=$Execute, Force=$Force)"
        Sync-DLToTeam -DLDisplayName $dlAddress -TeamDisplayName $TeamDisplayName -Execute:$Execute -Force:$Force
    }

    $organizationFile = Get-CapwatchOrganizationPath
    $organizationRows = @(Import-CapwatchOrganizationRows -Path $organizationFile)
    $commanderRows = @(Import-CapwatchCommanderRows)
    $cadetDutyPositionRows = @(Import-CapwatchCadetDutyPositionRows)
    $dutyPositionRows = @(Import-CapwatchDutyPositionRows)
    $teamGroups = @(Get-AllTeamGroups)
    Write-Log "Discovered $($teamGroups.Count) Microsoft 365 groups with Teams enabled."

    $manualGuestPath = Get-ManualGuestsCsvPath
    $manualGuests = @(Import-ManualGuestRows -Path $manualGuestPath)
    if ($manualGuests.Count -gt 0) {
        Write-Log "Loaded $($manualGuests.Count) manual guest record(s) for Team assignment from $manualGuestPath"
    }
    Sync-ManualGuestsToTeams -ManualGuests $manualGuests -TeamGroups $teamGroups -Execute:$execute

    $cacTeamName = Get-CacTeamDisplayName
    Sync-CacTeamFromCapwatch -CadetDutyPositions $cadetDutyPositionRows -DutyPositions $dutyPositionRows -TeamGroups $teamGroups -Execute:$execute
    $processedTeamNames[$cacTeamName.ToLowerInvariant()] = $true

    $teamOrganizations = @(
        Get-CapwatchActiveWingOrganizations -OrganizationRows $organizationRows |
            Where-Object {
                ($_.Scope -eq "UNIT" -or $_.Scope -eq "GROUP") -and
                -not (Test-CapwatchAdministrativeUnitCode -Unit $_.Unit)
            } |
            Sort-Object Unit -Unique
    )

    Write-Log "Discovered $($teamOrganizations.Count) active CAPWATCH unit/group organization(s) for Team sync."

    foreach ($organization in $teamOrganizations) {
        $teamsForOrganization = @(Find-TeamGroupsForOrganization -Organization $organization -TeamGroups $teamGroups)
        $prefix = Get-OrganizationTeamPrefix -Organization $organization
        $createdTeamForOrganization = $false

        if ($teamsForOrganization.Count -eq 0) {
            Write-Log "No Team found with prefix '$prefix' for active CAPWATCH organization '$($organization.Name)' ($($organization.Type)/$($organization.Scope))."

            if (-not $createMissingTeams) {
                Write-Log "Skipping Team creation for '$prefix' because CREATE_MISSING_TEAMS is not enabled."
                continue
            }

            if (-not (Test-TeamCreationUnitIncluded -Organization $organization -UnitFilter $createMissingTeamUnits)) {
                Write-Log "Skipping Team creation for '$prefix' because it is not included in CREATE_MISSING_TEAM_UNITS."
                continue
            }

            $createdTeam = New-CapwatchTeamForOrganization -Organization $organization -Commanders $commanderRows -Execute:$execute
            if (-not $createdTeam) {
                continue
            }

            $teamsForOrganization = @($createdTeam)
            $teamGroups += @($createdTeam)
            $createdTeamForOrganization = $true
        }

        foreach ($team in $teamsForOrganization) {
            $teamKey = $team.displayName.ToLowerInvariant()
            if ($processedTeamNames.ContainsKey($teamKey)) {
                Write-Log "Skipping Team '$($team.displayName)' because it was already processed by explicit mapping."
                continue
            }

            if ($createdTeamForOrganization) {
                Write-Log "New Team '$($team.displayName)' was created; syncing members from '$prefix@cowg.cap.gov' now."
            }

            Sync-FLToTeam -TeamDisplayName $team.displayName -Execute:$execute -Force:$force
            $processedTeamNames[$teamKey] = $true
        }
    }



    Write-Log "updateTeams timer function completed"
} catch {
    Write-Log "Unhandled exception in updateTeams function: $_"
    if ($_.Exception) { Write-Log $_.Exception.ToString() }
    throw
}
