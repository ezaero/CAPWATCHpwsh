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

    # Find team by display name
    try {
        $team = Get-MgGroup -Filter "displayName eq '$TeamDisplayName'" -ConsistencyLevel eventual
    } catch {
        $team = $null
    }
    if (-not $team) {
        $teams = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$TeamDisplayName'&`$select=id,displayName" -Headers $headers
        $team = $teams.value | Select-Object -First 1
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

    # Dry-run: log number of members that would be added
    Write-Log "Dry-run: $($toAdd.count) members to add for Team '$TeamDisplayName'"
    if ($toAdd.Count -gt 0) {
        foreach ($u in $toAdd[0..([Math]::Min(9,$toAdd.Count-1))]) {
            Write-Log ("Candidate: {0} <{1}> ({2})" -f $u.displayName, $u.mail, $u.userPrincipalName)
        }
        if ($toAdd.Count -gt 10) { Write-Log "...and $($toAdd.Count - 10) more candidates not shown" }
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
    } else {
        Write-Log "Dry-run complete for Team '$TeamDisplayName'. Re-run with EXECUTE=true to actually add members."
    }
}

### Main function entry (timer-triggered)
Write-Log "updateTeams timer function invoked"

try {
    # Determine execution flags from environment
    # Default to dry-run (no writes) because this runs as an Azure Function (non-interactive).
    $execute = $false
    $force = $true   # default to true so non-interactive hosts won't prompt
    if ($env:EXECUTE) { $execute = $env:EXECUTE.ToLower() -eq 'true' }
    if ($env:FORCE) { $force = $env:FORCE.ToLower() -eq 'true' }
    Write-Log ("Execution flags: EXECUTE={0}, FORCE={1} (set these as Function App settings to change behavior)" -f $execute, $force)

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
        Write-Log "No group->team mappings found (expected $mappingFile or env GROUP_TEAM_PAIRS). Exiting."
    }

    foreach ($map in $mappings) {
        if ($null -eq $map.GroupId -or $null -eq $map.TeamDisplayName) { Write-Log "Skipping invalid mapping entry: $map"; continue }
        Sync-GroupToTeam -GroupId $map.GroupId -TeamDisplayName $map.TeamDisplayName -Execute:$execute -Force:$force
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

    # Batch: call Sync-FLToTeam for each team provided
    Sync-FLToTeam -TeamDisplayName 'CO-022 Vance Brand Cadet' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-030 WolfPack' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-053 Eagle county Composite' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-068 NORTH VALLEY COMPOSITE SQDN TEAM' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-072 Boulder Composite' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-080 Pikes Peak Composite Squadron' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-099 Broomfield Composite Squadron' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-136 Jefferson County Senior Squadron' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-141 Montrose Composite' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-143 Mile High Cadet' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-147 Thompson Valley Composite' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-148 Mustang Squadron' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-157  Castle Rock Cadet Squadron' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-159 Air Academy' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-159 AIR ACADEMY CADET SQUADRON TEAM' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-162 Black Sheep Senior Squadron' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-163 HIGHLANDER COMPOSITE SQUADRON TEAM' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-164 GROUP 4 HEADQUARTERS TEAM' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-164 Group 4 Hq' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-165 GROUP 3 HEADQUARTERS TEAM' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-165 Group 3 Hq' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-167 GROUP 1 HEADQUARTERS TEAM' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-169 GROUP 2 HEADQUARTERS TEAM' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-169 Group 2 Hq' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-173 PARKER COMPOSITE SQDN TEAM' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-181 Steamboat Springs Composite Squadron' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-183 Valkyrie Cadet' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-183 VALKYRIE CADET SQDN TEAM' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-186 Dakota Ridge Composite' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-189 Mesa Verde Composite' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-189 MESA VERDE COMPOSITE SQUADRON TEAM' -Execute:$execute -Force:$force
    Sync-FLToTeam -TeamDisplayName 'CO-191 PLATTE VALLEY CADET SQUADRON TEAM' -Execute:$execute -Force:$force



    Write-Log "updateTeams timer function completed"
} catch {
    Write-Log "Unhandled exception in updateTeams function: $_"
    if ($_.Exception) { Write-Log $_.Exception.ToString() }
    throw
}