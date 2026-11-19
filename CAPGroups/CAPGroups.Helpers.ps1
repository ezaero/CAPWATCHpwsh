. "$PSScriptRoot\..\shared\OpsQuals.Helpers.ps1"

$script:CAPGroupsGraphRoot = "https://graph.microsoft.com/v1.0"
$script:CAPOpsQualAchievementIds = @(
    "44", "53", "55", "56", "57", "58", "59", "60", "61", "63", "64", "65", "66", "67", "68", "69",
    "70", "71", "72", "73", "74", "75", "76", "77", "78", "79", "80", "81", "90", "91", "96", "103",
    "124", "125", "126", "127", "128", "131", "132", "133", "134", "138", "145", "152", "164", "166",
    "184", "186", "189", "190", "191", "193", "195", "196", "211", "214", "215", "217", "218", "219",
    "220", "221", "222", "223", "231", "232", "233", "234", "235", "238", "246", "247", "248", "249",
    "250", "253", "254", "255", "256", "257", "258", "259", "260", "261", "262", "263", "264", "265",
    "269", "280", "284", "285", "286", "287", "292", "293", "294", "295"
)

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

function Get-CAPUserIdsByCapIds {
    param (
        [array]$CapIds,
        [hashtable]$UserIdByCapId
    )

    return @(
        foreach ($capId in @($CapIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)) {
            $key = "$capId".Trim()
            if ($UserIdByCapId.ContainsKey($key)) {
                $UserIdByCapId[$key]
            }
        }
    ) | Sort-Object -Unique
}

function Get-CAPAchievementLookupById {
    param (
        [array]$AchievementLookupRows
    )

    $lookup = @{}
    foreach ($achievement in @($AchievementLookupRows)) {
        if ($null -eq $achievement -or [string]::IsNullOrWhiteSpace($achievement.AchvID)) {
            continue
        }

        $lookup["$($achievement.AchvID)".Trim()] = $achievement
    }

    return $lookup
}

function Get-CAPUserIdsByCapIdsAndCompanyNames {
    param (
        [array]$CapIds,
        [array]$CompanyNames,
        [array]$AllUsers
    )

    $capIdSet = @{}
    foreach ($capId in @($CapIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        $capIdSet["$capId".Trim()] = $true
    }

    $companyNameSet = @{}
    foreach ($companyName in @($CompanyNames | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        $companyNameSet["$companyName".Trim().ToUpperInvariant()] = $true
    }

    return @(
        $AllUsers |
            Where-Object {
                -not [string]::IsNullOrWhiteSpace($_.id) -and
                -not [string]::IsNullOrWhiteSpace($_.employeeId) -and
                $capIdSet.ContainsKey("$($_.employeeId)".Trim()) -and
                $companyNameSet.ContainsKey("$($_.companyName)".Trim().ToUpperInvariant())
            } |
            Select-Object -ExpandProperty id
    ) | Sort-Object -Unique
}

function Test-CAPUserMatchesAudience {
    param (
        [object]$User,
        [ValidateSet("All", "Seniors", "Cadets")]
        [string]$Audience
    )

    if ($Audience -eq "All") {
        return $true
    }

    $employeeType = "$($User.employeeType)".Trim().ToUpperInvariant()
    $department = "$($User.department)".Trim().ToUpperInvariant()

    if ($Audience -eq "Seniors") {
        return $employeeType -eq "SENIOR"
    }

    return $employeeType -eq "CADET" -or $employeeType -eq "PARENT" -or $department -like "*CP*" -or $department -like "*EX*"
}

function Get-CAPUserIdsByCompanyNamesAndAudience {
    param (
        [array]$CompanyNames,
        [array]$AllUsers,
        [ValidateSet("All", "Seniors", "Cadets")]
        [string]$Audience = "All"
    )

    $companyNameSet = @{}
    foreach ($companyName in @($CompanyNames | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        $companyNameSet["$companyName".Trim().ToUpperInvariant()] = $true
    }

    return @(
        $AllUsers |
            Where-Object {
                -not [string]::IsNullOrWhiteSpace($_.id) -and
                $companyNameSet.ContainsKey("$($_.companyName)".Trim().ToUpperInvariant()) -and
                (Test-CAPUserMatchesAudience -User $_ -Audience $Audience)
            } |
            Select-Object -ExpandProperty id
    ) | Sort-Object -Unique
}

function Test-CAPAdministrativeUnitCode {
    param (
        [string]$Unit
    )

    return $Unit -in @("000", "001", "999")
}

function Get-CAPActiveGroupUnitCodes {
    param (
        [array]$OrganizationRows,
        [Parameter(Mandatory = $true)]
        [string]$GroupNumber,
        [string]$WingDesignator
    )

    $activeOrganizations = @(
        $OrganizationRows |
            Where-Object {
                "$($_.Wing)".Trim().ToUpperInvariant() -eq $WingDesignator -and
                "$($_.Status)".Trim().ToUpperInvariant() -eq "ACTIVE"
            }
    )
    $group = $activeOrganizations |
        Where-Object { ($_.Scope -eq "GROUP" -or $_.Type -eq "GROUP") -and $_.Name -match "(?i)\bGROUP\s+$GroupNumber\b" } |
        Select-Object -First 1

    if (-not $group) {
        return @()
    }

    $unitCodes = @(
        $activeOrganizations |
            Where-Object {
                $_.Scope -eq "UNIT" -and
                $_.NextLevel -eq $group.ORGID -and
                -not (Test-CAPAdministrativeUnitCode -Unit $_.Unit)
            } |
            Select-Object -ExpandProperty Unit
    )

    if (-not [string]::IsNullOrWhiteSpace($group.Unit)) {
        $unitCodes += $group.Unit
    }

    return @($unitCodes | Sort-Object -Unique)
}

function New-CAPAudienceSecurityGroupPlans {
    param (
        [array]$OrganizationRows,
        [array]$AllUsers,
        [string]$WingDesignator
    )

    $activeOrganizations = @(
        $OrganizationRows |
            Where-Object {
                "$($_.Wing)".Trim().ToUpperInvariant() -eq $WingDesignator -and
                "$($_.Status)".Trim().ToUpperInvariant() -eq "ACTIVE"
            }
    )
    $activeUnits = @(
        $activeOrganizations |
            Where-Object { $_.Scope -eq "UNIT" -and -not (Test-CAPAdministrativeUnitCode -Unit $_.Unit) } |
            Sort-Object Unit
    )
    $activeGroups = @(
        $activeOrganizations |
            Where-Object { $_.Scope -eq "GROUP" -or $_.Type -eq "GROUP" } |
            Sort-Object Name
    )

    $plans = @()
    $wingCompanyNames = @($activeUnits | ForEach-Object { "$WingDesignator-$($_.Unit)" })
    foreach ($audience in @("Seniors", "Cadets")) {
        $plans += New-CAPSecurityGroupPlan `
            -DisplayName "CAP-WingAudience-$WingDesignator-$audience" `
            -Description "Entra audience group for $WingDesignator Wing $audience distribution lists." `
            -Category "Audience" `
            -MemberIds (Get-CAPUserIdsByCompanyNamesAndAudience -CompanyNames $wingCompanyNames -AllUsers $AllUsers -Audience $audience)
    }

    foreach ($group in $activeGroups) {
        if ($group.Name -notmatch '(?i)\bGROUP\s+(\d+)\b') {
            continue
        }

        $groupNumber = $matches[1]
        $unitCodes = @(
            $activeUnits |
                Where-Object { $_.NextLevel -eq $group.ORGID } |
                Select-Object -ExpandProperty Unit
        )
        if ($unitCodes.Count -eq 0) {
            continue
        }

        $plans += New-CAPSecurityGroupPlan `
            -DisplayName "CAP-GroupAudience-$WingDesignator-$groupNumber-All" `
            -Description "Entra audience group for $WingDesignator Group $groupNumber distribution lists." `
            -Category "Audience" `
            -MemberIds (Get-CAPUserIdsByCompanyNamesAndAudience -CompanyNames @($unitCodes | ForEach-Object { "$WingDesignator-$_" }) -AllUsers $AllUsers -Audience "All")
    }

    foreach ($unit in $activeUnits) {
        $companyName = "$WingDesignator-$($unit.Unit)"
        foreach ($audience in @("All", "Seniors", "Cadets")) {
            $plans += New-CAPSecurityGroupPlan `
                -DisplayName "CAP-UnitAudience-$WingDesignator-$($unit.Unit)-$audience" `
                -Description "Entra audience group for $companyName $audience distribution lists." `
                -Category "Audience" `
                -MemberIds (Get-CAPUserIdsByCompanyNamesAndAudience -CompanyNames @($companyName) -AllUsers $AllUsers -Audience $audience)
        }
    }

    return @($plans)
}

function New-CAPOpsQualSecurityGroupPlans {
    param (
        [array]$AchievementRows,
        [array]$AchievementLookupRows = @(),
        [array]$TaskRows,
        [array]$OrganizationRows,
        [array]$AllUsers,
        [hashtable]$UserIdByCapId,
        [string]$WingDesignator,
        [array]$OpsQualAchievementIds = $script:CAPOpsQualAchievementIds,
        [array]$KapaUnitCodes = @('162', '148', '157', '163', '164', '183', '186', '143', '031'),
        [array]$ManualKapaCapIds = @('446885', '344160', '574341', '100137')
    )

    $activePilotCapIds = @(Get-CAPPilotCapIds -TaskRows $TaskRows)
    $esCapIds = @(Get-CAPAchievementCapIds -AchievementRows $AchievementRows -AchievementIds @("53") -Statuses @("ACTIVE"))
    $missionPilotCapIds = @(
        Get-CAPAchievementCapIds -AchievementRows $AchievementRows -AchievementIds @("57") -Statuses @("ACTIVE", "TRAINING") |
            Where-Object { $activePilotCapIds -contains $_ }
    )
    $activeMissionPilotCapIds = @(
        Get-CAPAchievementCapIds -AchievementRows $AchievementRows -AchievementIds @("57") -Statuses @("ACTIVE") |
            Where-Object { $activePilotCapIds -contains $_ }
    )
    $mfcCapIds = @(Get-CAPAchievementCapIds -AchievementRows $AchievementRows -AchievementIds @("103") -Statuses @("ACTIVE"))

    $kapaCapIds = @(Get-CAPPilotCapIds -TaskRows $TaskRows -ManualCapIds $ManualKapaCapIds)
    $kapaCompanyNames = @($KapaUnitCodes | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { "$WingDesignator-$($_.Trim())" })
    $group4UnitCodes = @(Get-CAPActiveGroupUnitCodes -OrganizationRows $OrganizationRows -GroupNumber "4" -WingDesignator $WingDesignator)
    if ($group4UnitCodes.Count -eq 0) {
        $group4UnitCodes = @('157', '162', '163', '173', '143', '148', '186', '183', '164', '031')
    }
    $group4CompanyNames = @($group4UnitCodes | ForEach-Object { "$WingDesignator-$($_.Trim())" })

    $exactAchievementPlans = @()
    $achievementLookupById = Get-CAPAchievementLookupById -AchievementLookupRows $AchievementLookupRows
    if ($achievementLookupById.Count -gt 0) {
        foreach ($achievementId in @($OpsQualAchievementIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)) {
            $normalizedAchievementId = "$achievementId".Trim()
            if (-not $achievementLookupById.ContainsKey($normalizedAchievementId)) {
                Write-Warning "Skipping OpsQual AchvID '$normalizedAchievementId' because it was not found in Achievements.txt."
                continue
            }

            $achievementName = "$($achievementLookupById[$normalizedAchievementId].Achv)".Trim()
            $achievementNamePiece = ConvertTo-CAPGroupNamePiece -Value $achievementName
            if ([string]::IsNullOrWhiteSpace($achievementNamePiece)) {
                Write-Warning "Skipping OpsQual AchvID '$normalizedAchievementId' because its achievement name is empty."
                continue
            }

            $exactAchievementPlans += New-CAPSecurityGroupPlan `
                -DisplayName "CAP-OpsQual-$achievementNamePiece" `
                -Description "CAPWATCH active or training $achievementName qualification (AchvID $normalizedAchievementId)." `
                -Category "OpsQual" `
                -MemberIds (Get-CAPUserIdsByCapIds -CapIds (Get-CAPAchievementCapIds -AchievementRows $AchievementRows -AchievementIds @($normalizedAchievementId) -Statuses @("ACTIVE", "TRAINING")) -UserIdByCapId $UserIdByCapId) `
                -Metadata @{
                    AchievementId = $normalizedAchievementId
                    AchievementName = $achievementName
                    Statuses = @("ACTIVE", "TRAINING")
                }
        }
    }

    $definitions = @(
        [PSCustomObject]@{ Name = "CAP-OpsQual-Pilot"; Description = "CAPWATCH active pilots with TaskID 69."; MemberIds = @(Get-CAPUserIdsByCapIds -CapIds $activePilotCapIds -UserIdByCapId $UserIdByCapId) },
        [PSCustomObject]@{ Name = "CAP-OpsQual-KAPA-Pilot"; Description = "CAPWATCH active KAPA-scope pilots plus approved manual CAPIDs."; MemberIds = @(Get-CAPUserIdsByCapIdsAndCompanyNames -CapIds $kapaCapIds -CompanyNames $kapaCompanyNames -AllUsers $AllUsers) },
        [PSCustomObject]@{ Name = "CAP-OpsQual-ES"; Description = "CAPWATCH active emergency services qualified members with AchvID 53."; MemberIds = @(Get-CAPUserIdsByCapIds -CapIds $esCapIds -UserIdByCapId $UserIdByCapId) },
        [PSCustomObject]@{ Name = "CAP-OpsQual-ES-Group-4"; Description = "CAPWATCH active emergency services qualified members in Group 4 units."; MemberIds = @(Get-CAPUserIdsByCapIdsAndCompanyNames -CapIds $esCapIds -CompanyNames $group4CompanyNames -AllUsers $AllUsers) },
        [PSCustomObject]@{ Name = "CAP-OpsQual-Mission-Check-Pilot"; Description = "CAPWATCH active mission check pilots with AchvID 90."; MemberIds = @(Get-CAPUserIdsByCapIds -CapIds (Get-CAPAchievementCapIds -AchievementRows $AchievementRows -AchievementIds @("90") -Statuses @("ACTIVE")) -UserIdByCapId $UserIdByCapId) },
        [PSCustomObject]@{ Name = "CAP-OpsQual-Mission-Check-Pilot-Examiner"; Description = "CAPWATCH active mission check pilot examiners with AchvID 166."; MemberIds = @(Get-CAPUserIdsByCapIds -CapIds (Get-CAPAchievementCapIds -AchievementRows $AchievementRows -AchievementIds @("166") -Statuses @("ACTIVE")) -UserIdByCapId $UserIdByCapId) },
        [PSCustomObject]@{ Name = "CAP-OpsQual-Aircrew"; Description = "CAPWATCH active or training aircrew with AchvID 55."; MemberIds = @(Get-CAPUserIdsByCapIds -CapIds (Get-CAPAchievementCapIds -AchievementRows $AchievementRows -AchievementIds @("55") -Statuses @("ACTIVE", "TRAINING")) -UserIdByCapId $UserIdByCapId) },
        [PSCustomObject]@{ Name = "CAP-OpsQual-Mission-Pilot"; Description = "CAPWATCH active or training mission pilots who are also active pilots."; MemberIds = @(Get-CAPUserIdsByCapIds -CapIds $missionPilotCapIds -UserIdByCapId $UserIdByCapId) },
        [PSCustomObject]@{ Name = "CAP-OpsQual-Mission-Pilot-MFC"; Description = "CAPWATCH active mission pilots with active Mountain Flying Certification."; MemberIds = @(Get-CAPUserIdsByCapIds -CapIds (@($activeMissionPilotCapIds | Where-Object { $mfcCapIds -contains $_ })) -UserIdByCapId $UserIdByCapId) },
        [PSCustomObject]@{ Name = "CAP-OpsQual-Instructor-Pilot"; Description = "CAPWATCH active or training instructor pilots with AchvID 59."; MemberIds = @(Get-CAPUserIdsByCapIds -CapIds (Get-CAPAchievementCapIds -AchievementRows $AchievementRows -AchievementIds @("59") -Statuses @("ACTIVE", "TRAINING")) -UserIdByCapId $UserIdByCapId) },
        [PSCustomObject]@{ Name = "CAP-OpsQual-Check-Pilot-Examiner"; Description = "CAPWATCH active check pilot examiners with AchvID 164."; MemberIds = @(Get-CAPUserIdsByCapIds -CapIds (Get-CAPAchievementCapIds -AchievementRows $AchievementRows -AchievementIds @("164") -Statuses @("ACTIVE")) -UserIdByCapId $UserIdByCapId) },
        [PSCustomObject]@{ Name = "CAP-OpsQual-Orientation-Pilot"; Description = "CAPWATCH active or training orientation pilots with AchvID 91."; MemberIds = @(Get-CAPUserIdsByCapIds -CapIds (Get-CAPAchievementCapIds -AchievementRows $AchievementRows -AchievementIds @("91") -Statuses @("ACTIVE", "TRAINING")) -UserIdByCapId $UserIdByCapId) },
        [PSCustomObject]@{ Name = "CAP-OpsQual-Communicator"; Description = "CAPWATCH active or training communicators with AchvID 217."; MemberIds = @(Get-CAPUserIdsByCapIds -CapIds (Get-CAPAchievementCapIds -AchievementRows $AchievementRows -AchievementIds @("217") -Statuses @("ACTIVE", "TRAINING")) -UserIdByCapId $UserIdByCapId) },
        [PSCustomObject]@{ Name = "CAP-OpsQual-Mission-Base-Staff"; Description = "CAPWATCH active or training mission base staff qualifications."; MemberIds = @(Get-CAPUserIdsByCapIds -CapIds (Get-CAPAchievementCapIds -AchievementRows $AchievementRows -AchievementIds @("61", "63", "64", "65", "66", "67", "68", "75", "76", "77", "78", "79", "80") -Statuses @("ACTIVE", "TRAINING")) -UserIdByCapId $UserIdByCapId) },
        [PSCustomObject]@{ Name = "CAP-OpsQual-sUAS"; Description = "CAPWATCH active or training sUAS qualifications."; MemberIds = @(Get-CAPUserIdsByCapIds -CapIds (Get-CAPAchievementCapIds -AchievementRows $AchievementRows -AchievementIds @("257", "258", "262", "263") -Statuses @("ACTIVE", "TRAINING")) -UserIdByCapId $UserIdByCapId) }
    )

    $rollupPlans = @(
        foreach ($definition in $definitions) {
            New-CAPSecurityGroupPlan `
                -DisplayName $definition.Name `
                -Description $definition.Description `
                -Category "OpsQual" `
                -MemberIds $definition.MemberIds
        }
    )

    $seenDisplayNames = @{}
    return @(
        foreach ($plan in @($exactAchievementPlans + $rollupPlans)) {
            if ($seenDisplayNames.ContainsKey($plan.DisplayName)) {
                continue
            }

            $seenDisplayNames[$plan.DisplayName] = $true
            $plan
        }
    )
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

        [array]$AchievementRows = @(),

        [array]$AchievementLookupRows = @(),

        [array]$OpsQualAchievementIds = $script:CAPOpsQualAchievementIds,

        [array]$TaskRows = @(),

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

    $plans += @(New-CAPAudienceSecurityGroupPlans `
        -OrganizationRows $OrganizationRows `
        -AllUsers $AllUsers `
        -WingDesignator $wingDesignator)

    if ($AchievementRows.Count -gt 0 -or $TaskRows.Count -gt 0) {
        $plans += @(New-CAPOpsQualSecurityGroupPlans `
            -AchievementRows $AchievementRows `
            -AchievementLookupRows $AchievementLookupRows `
            -OpsQualAchievementIds $OpsQualAchievementIds `
            -TaskRows $TaskRows `
            -OrganizationRows $OrganizationRows `
            -AllUsers $AllUsers `
            -UserIdByCapId $userIdByCapId `
            -WingDesignator $wingDesignator)
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

