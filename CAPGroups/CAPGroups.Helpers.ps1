$script:CAPGroupsGraphRoot = "https://graph.microsoft.com/v1.0"

function ConvertTo-CAPGroupNamePiece {
    param (
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $words = [regex]::Matches($Value, "[A-Za-z0-9]+") | ForEach-Object { $_.Value }
    if ($words.Count -eq 0) {
        return $null
    }

    return ($words -join "-")
}

function ConvertTo-CAPGroupMailNickname {
    param (
        [Parameter(Mandatory = $true)]
        [string]$DisplayName
    )

    $nickname = $DisplayName -replace "[^A-Za-z0-9]", ""
    if ($nickname.Length -gt 64) {
        return $nickname.Substring(0, 64)
    }

    return $nickname
}

function New-CAPSecurityGroupPlan {
    param (
        [Parameter(Mandatory = $true)]
        [string]$DisplayName,

        [array]$MemberIds = @(),

        [string]$Description = "",

        [string]$Category = "CAPWATCH",

        [hashtable]$Metadata = @{}
    )

    $memberIds = @(
        $MemberIds |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { "$_".Trim() } |
            Sort-Object -Unique
    )

    if ([string]::IsNullOrWhiteSpace($Description)) {
        $Description = "CAPWATCH-managed security group."
    }

    return [PSCustomObject]@{
        DisplayName  = $DisplayName
        MailNickname = ConvertTo-CAPGroupMailNickname -DisplayName $DisplayName
        Description  = $Description
        Category     = $Category
        MemberIds    = $memberIds
        Metadata     = $Metadata
    }
}

function Get-CAPUserIdByCapIdMap {
    param (
        [array]$AllUsers
    )

    $map = @{}
    foreach ($user in $AllUsers) {
        if ($null -eq $user -or [string]::IsNullOrWhiteSpace($user.employeeId) -or [string]::IsNullOrWhiteSpace($user.id)) {
            continue
        }

        $capId = "$($user.employeeId)".Trim()
        if (-not $map.ContainsKey($capId)) {
            $map[$capId] = "$($user.id)".Trim()
        }
    }

    return $map
}

function Get-CAPOrganizationByIdMap {
    param (
        [array]$OrganizationRows
    )

    $map = @{}
    foreach ($org in $OrganizationRows) {
        if ($null -eq $org -or [string]::IsNullOrWhiteSpace($org.ORGID)) {
            continue
        }

        $map["$($org.ORGID)".Trim()] = $org
    }

    return $map
}

function Get-CAPMemberUserIds {
    param (
        [array]$MemberRows,
        [hashtable]$UserIdByCapId,
        [scriptblock]$Predicate
    )

    return @(
        foreach ($member in $MemberRows) {
            if ($null -eq $member -or [string]::IsNullOrWhiteSpace($member.CAPID)) {
                continue
            }

            if ($Predicate -and -not (& $Predicate $member)) {
                continue
            }

            $capId = "$($member.CAPID)".Trim()
            if ($UserIdByCapId.ContainsKey($capId)) {
                $UserIdByCapId[$capId]
            }
        }
    ) | Sort-Object -Unique
}

function Get-CAPDutyGroupDisplayName {
    param (
        [Parameter(Mandatory = $true)]
        [object]$DutyPosition,

        [object]$Organization,

        [object]$Member
    )

    $level = ConvertTo-CAPGroupNamePiece -Value $DutyPosition.Lvl
    if ([string]::IsNullOrWhiteSpace($level)) {
        $level = "UNKNOWN"
    }

    $wing = ConvertTo-CAPGroupNamePiece -Value $Organization.Wing
    if ([string]::IsNullOrWhiteSpace($wing) -and $null -ne $Member) {
        $wing = ConvertTo-CAPGroupNamePiece -Value $Member.Wing
    }

    $orgKey = $wing
    if ([string]::IsNullOrWhiteSpace($orgKey)) {
        $orgKey = ConvertTo-CAPGroupNamePiece -Value $DutyPosition.ORGID
    }
    if ([string]::IsNullOrWhiteSpace($orgKey)) {
        $orgKey = "unknown"
    }

    $functionalArea = ConvertTo-CAPGroupNamePiece -Value $DutyPosition.FunctArea
    $assistant = $null
    if ("$($DutyPosition.Asst)".Trim() -eq "1") {
        $assistant = "Assistant"
    }

    $dutyName = ConvertTo-CAPGroupNamePiece -Value $DutyPosition.Duty
    if ([string]::IsNullOrWhiteSpace($dutyName)) {
        $dutyName = "Duty"
    }

    return (@("CAP", "Duty", $level, $orgKey, $functionalArea, $assistant, $dutyName) |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join "-"
}

function New-CAPSecurityGroupPlans {
    param (
        [Parameter(Mandatory = $true)]
        [array]$MemberRows,

        [Parameter(Mandatory = $true)]
        [array]$DutyPositionRows,

        [Parameter(Mandatory = $true)]
        [array]$OrganizationRows,

        [Parameter(Mandatory = $true)]
        [array]$AllUsers,

        [string]$WingDesignator = $env:WING_DESIGNATOR
    )

    if ([string]::IsNullOrWhiteSpace($WingDesignator)) {
        $WingDesignator = "CO"
    }

    $wingDesignator = $WingDesignator.Trim().ToUpperInvariant()
    $userIdByCapId = Get-CAPUserIdByCapIdMap -AllUsers $AllUsers
    $orgById = Get-CAPOrganizationByIdMap -OrganizationRows $OrganizationRows
    $memberByCapId = @{}
    foreach ($member in $MemberRows) {
        if ($member.CAPID) {
            $memberByCapId["$($member.CAPID)".Trim()] = $member
        }
    }

    $wingMembers = @($MemberRows | Where-Object { "$($_.Wing)".Trim().ToUpperInvariant() -eq $wingDesignator })
    $activeWingMembers = @($wingMembers | Where-Object { "$($_.MbrStatus)".Trim().ToUpperInvariant() -eq "ACTIVE" })
    $plans = @()

    $plans += New-CAPSecurityGroupPlan `
        -DisplayName "CAP-Wing-$wingDesignator" `
        -Description "CAPWATCH active members in $wingDesignator Wing." `
        -Category "Wing" `
        -MemberIds (Get-CAPMemberUserIds -MemberRows $activeWingMembers -UserIdByCapId $userIdByCapId)

    foreach ($region in @($activeWingMembers | Select-Object -ExpandProperty Region -Unique | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object)) {
        $regionKey = ConvertTo-CAPGroupNamePiece -Value $region
        $plans += New-CAPSecurityGroupPlan `
            -DisplayName "CAP-Region-$regionKey" `
            -Description "CAPWATCH active members in $region region." `
            -Category "Region" `
            -MemberIds (Get-CAPMemberUserIds -MemberRows $activeWingMembers -UserIdByCapId $userIdByCapId -Predicate { param($m) "$($m.Region)".Trim() -eq $region })
    }

    foreach ($unit in @($activeWingMembers | Select-Object -ExpandProperty Unit -Unique | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object)) {
        $unitKey = ConvertTo-CAPGroupNamePiece -Value $unit
        $plans += New-CAPSecurityGroupPlan `
            -DisplayName "CAP-Unit-$wingDesignator-$unitKey" `
            -Description "CAPWATCH active members in $wingDesignator-$unit." `
            -Category "Unit" `
            -MemberIds (Get-CAPMemberUserIds -MemberRows $activeWingMembers -UserIdByCapId $userIdByCapId -Predicate { param($m) "$($m.Unit)".Trim() -eq $unit })
    }

    foreach ($type in @($activeWingMembers | Select-Object -ExpandProperty Type -Unique | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object)) {
        $typeKey = ConvertTo-CAPGroupNamePiece -Value $type
        $plans += New-CAPSecurityGroupPlan `
            -DisplayName "CAP-MemberType-$typeKey" `
            -Description "CAPWATCH active $type members in $wingDesignator Wing." `
            -Category "MemberType" `
            -MemberIds (Get-CAPMemberUserIds -MemberRows $activeWingMembers -UserIdByCapId $userIdByCapId -Predicate { param($m) "$($m.Type)".Trim() -eq $type })
    }

    foreach ($status in @($wingMembers | Select-Object -ExpandProperty MbrStatus -Unique | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object)) {
        $statusKey = ConvertTo-CAPGroupNamePiece -Value $status
        $plans += New-CAPSecurityGroupPlan `
            -DisplayName "CAP-MemberStatus-$statusKey" `
            -Description "CAPWATCH $status members in $wingDesignator Wing." `
            -Category "MemberStatus" `
            -MemberIds (Get-CAPMemberUserIds -MemberRows $wingMembers -UserIdByCapId $userIdByCapId -Predicate { param($m) "$($m.MbrStatus)".Trim() -eq $status })
    }

    $dutyGroups = @{}
    foreach ($duty in $DutyPositionRows) {
        if ($null -eq $duty -or [string]::IsNullOrWhiteSpace($duty.CAPID)) {
            continue
        }

        $capId = "$($duty.CAPID)".Trim()
        if (-not $userIdByCapId.ContainsKey($capId)) {
            continue
        }

        $org = $null
        if (-not [string]::IsNullOrWhiteSpace($duty.ORGID) -and $orgById.ContainsKey("$($duty.ORGID)".Trim())) {
            $org = $orgById["$($duty.ORGID)".Trim()]
        }
        $member = $memberByCapId[$capId]
        $resolvedWing = "$($org.Wing)".Trim()
        if ([string]::IsNullOrWhiteSpace($resolvedWing) -and $null -ne $member) {
            $resolvedWing = "$($member.Wing)".Trim()
        }

        if ($resolvedWing.ToUpperInvariant() -ne $wingDesignator) {
            continue
        }

        $displayName = Get-CAPDutyGroupDisplayName -DutyPosition $duty -Organization $org -Member $member
        if (-not $dutyGroups.ContainsKey($displayName)) {
            $description = "CAPWATCH duty group for $($duty.Duty)"
            if ($org -and -not [string]::IsNullOrWhiteSpace($org.Name)) {
                $description = "$description at $($org.Name)"
            }

            $dutyGroups[$displayName] = [PSCustomObject]@{
                DisplayName = $displayName
                Description = $description
                MemberIds = @()
                Duty = $duty.Duty
                Level = $duty.Lvl
                FunctionalArea = $duty.FunctArea
                Assistant = ("$($duty.Asst)".Trim() -eq "1")
            }
        }

        $dutyGroups[$displayName].MemberIds += $userIdByCapId[$capId]
    }

    foreach ($group in @($dutyGroups.Values | Sort-Object DisplayName)) {
        $plans += New-CAPSecurityGroupPlan `
            -DisplayName $group.DisplayName `
            -Description $group.Description `
            -Category "Duty" `
            -MemberIds $group.MemberIds `
            -Metadata @{
                Duty = $group.Duty
                Level = $group.Level
                FunctionalArea = $group.FunctionalArea
                Assistant = $group.Assistant
            }
    }

    return @($plans | Sort-Object DisplayName)
}

function ConvertTo-CAPGraphFilterLiteral {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    return $Value.Replace("'", "''")
}

function Invoke-CAPGraphRequestWithRetry {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Method,

        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [object]$Body = $null,

        [int]$MaxAttempts = 6
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            if ($null -ne $Body) {
                return Invoke-MgGraphRequest -Method $Method -Uri $Uri -Body $Body -ErrorAction Stop
            }

            return Invoke-MgGraphRequest -Method $Method -Uri $Uri -ErrorAction Stop
        } catch {
            $statusCode = $null
            if ($null -ne $_.Exception.Response -and $null -ne $_.Exception.Response.StatusCode) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            $isRetryable = $statusCode -in @(404, 409, 429, 500, 502, 503, 504)
            if (-not $isRetryable -or $attempt -ge $MaxAttempts) {
                throw
            }

            $delaySeconds = [Math]::Min(60, [int](3 * [Math]::Pow(2, ($attempt - 1))))
            Write-Warning "Graph $Method $Uri returned $statusCode. Retrying in $delaySeconds seconds (attempt $attempt of $MaxAttempts)."
            Start-Sleep -Seconds $delaySeconds
        }
    }
}

function Get-CAPGraphSecurityGroupByDisplayName {
    param (
        [Parameter(Mandatory = $true)]
        [string]$DisplayName
    )

    $escaped = ConvertTo-CAPGraphFilterLiteral -Value $DisplayName
    $uri = "$script:CAPGroupsGraphRoot/groups?`$select=id,displayName,mailNickname,description,createdDateTime&`$filter=displayName eq '$escaped'"
    $response = Invoke-CAPGraphRequestWithRetry -Method GET -Uri $uri
    $groups = @($response.value)
    if ($groups.Count -gt 1) {
        $selectedGroup = @($groups | Sort-Object createdDateTime, id | Select-Object -First 1)[0]
        $duplicateIds = @($groups | Where-Object { $_.id -ne $selectedGroup.id } | Select-Object -ExpandProperty id)
        Write-Warning "Multiple Entra groups found with displayName '$DisplayName'. Using '$($selectedGroup.id)' and leaving duplicate group id(s): $($duplicateIds -join ', ')."
        return $selectedGroup
    }

    if ($groups.Count -eq 0) {
        return $null
    }

    return $groups[0]
}

function New-CAPGraphSecurityGroup {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Plan
    )

    $body = @{
        displayName = $Plan.DisplayName
        description = $Plan.Description
        mailEnabled = $false
        mailNickname = $Plan.MailNickname
        securityEnabled = $true
    }

    return Invoke-CAPGraphRequestWithRetry -Method POST -Uri "$script:CAPGroupsGraphRoot/groups" -Body $body
}

function Get-CAPGraphGroupMemberIds {
    param (
        [Parameter(Mandatory = $true)]
        [string]$GroupId
    )

    $memberIds = @()
    $uri = "$script:CAPGroupsGraphRoot/groups/$GroupId/members?`$select=id"
    while (-not [string]::IsNullOrWhiteSpace($uri)) {
        $response = Invoke-CAPGraphRequestWithRetry -Method GET -Uri $uri
        $memberIds += @($response.value | ForEach-Object { $_.id } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $uri = $response.'@odata.nextLink'
    }

    return @($memberIds | Sort-Object -Unique)
}

function Add-CAPGraphGroupMember {
    param (
        [Parameter(Mandatory = $true)]
        [string]$GroupId,

        [Parameter(Mandatory = $true)]
        [string]$UserId
    )

    $body = @{
        "@odata.id" = "$script:CAPGroupsGraphRoot/directoryObjects/$UserId"
    }
    try {
        Invoke-CAPGraphRequestWithRetry -Method POST -Uri "$script:CAPGroupsGraphRoot/groups/$GroupId/members/`$ref" -Body $body | Out-Null
    } catch {
        $message = $_.ToString()
        if ($message -match "added object references already exist" -or $message -match "Object references already exist") {
            Write-Warning "Graph reported user '$UserId' is already a member of group '$GroupId'. Treating as already reconciled."
            return
        }

        throw
    }
}

function Remove-CAPGraphGroupMember {
    param (
        [Parameter(Mandatory = $true)]
        [string]$GroupId,

        [Parameter(Mandatory = $true)]
        [string]$UserId
    )

    Invoke-CAPGraphRequestWithRetry -Method DELETE -Uri "$script:CAPGroupsGraphRoot/groups/$GroupId/members/$UserId/`$ref" | Out-Null
}

function Sync-CAPSecurityGroupPlan {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Plan,

        [bool]$Execute = $false
    )

    $group = Get-CAPGraphSecurityGroupByDisplayName -DisplayName $Plan.DisplayName
    $actions = @()

    if ($null -eq $group) {
        if ($Execute) {
            $group = New-CAPGraphSecurityGroup -Plan $Plan
            $actions += [PSCustomObject]@{ Action = "ensureGroup"; Status = "created"; DisplayName = $Plan.DisplayName; GroupId = $group.id }
        } else {
            $actions += [PSCustomObject]@{ Action = "ensureGroup"; Status = "wouldCreate"; DisplayName = $Plan.DisplayName; GroupId = $null }
            return [PSCustomObject]@{
                DisplayName = $Plan.DisplayName
                DesiredCount = @($Plan.MemberIds).Count
                AddCount = @($Plan.MemberIds).Count
                RemoveCount = 0
                Actions = $actions
            }
        }
    } else {
        $actions += [PSCustomObject]@{ Action = "ensureGroup"; Status = "exists"; DisplayName = $Plan.DisplayName; GroupId = $group.id }
    }

    $currentMemberIds = @(Get-CAPGraphGroupMemberIds -GroupId $group.id)
    $desiredMemberIds = @($Plan.MemberIds)
    $add = @($desiredMemberIds | Where-Object { $currentMemberIds -notcontains $_ })
    $remove = @($currentMemberIds | Where-Object { $desiredMemberIds -notcontains $_ })

    foreach ($memberId in $add) {
        if ($Execute) {
            Add-CAPGraphGroupMember -GroupId $group.id -UserId $memberId
            $status = "added"
        } else {
            $status = "wouldAdd"
        }
        $actions += [PSCustomObject]@{ Action = "ensureMember"; Status = $status; DisplayName = $Plan.DisplayName; GroupId = $group.id; UserId = $memberId }
    }

    foreach ($memberId in $remove) {
        if ($Execute) {
            Remove-CAPGraphGroupMember -GroupId $group.id -UserId $memberId
            $status = "removed"
        } else {
            $status = "wouldRemove"
        }
        $actions += [PSCustomObject]@{ Action = "removeStaleMember"; Status = $status; DisplayName = $Plan.DisplayName; GroupId = $group.id; UserId = $memberId }
    }

    return [PSCustomObject]@{
        DisplayName = $Plan.DisplayName
        DesiredCount = $desiredMemberIds.Count
        AddCount = $add.Count
        RemoveCount = $remove.Count
        Actions = $actions
    }
}

