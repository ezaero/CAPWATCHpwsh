function Get-RecruitingCapIds {
    param (
        [array]$specTracks,
        [array]$dutyPositions,
        [array]$commanders = @(),
        [string]$wingDesignator = $null
    )

    $trackCapIds = @(
        $specTracks |
            Where-Object { $_.Track -match '(?i)^RECRUITING AND RETENTION OFFICER$' } |
            Select-Object -ExpandProperty CAPID
    )

    $dutyCapIds = @(
        $dutyPositions |
            Where-Object {
                $_.Duty -match '(?i)^Recruiting Officer$' -or
                $_.FunctArea -match '(?i)^Recruiting Officer$' -or
                $_.Duty -match '(?i)(Commander|Chief of Staff)'
            } |
            Select-Object -ExpandProperty CAPID
    )

    $commanderCapIds = @(
        $commanders |
            Where-Object {
                -not [string]::IsNullOrWhiteSpace($_.CAPID) -and
                (
                    [string]::IsNullOrWhiteSpace($wingDesignator) -or
                    $_.Wing -eq $wingDesignator
                )
            } |
            Select-Object -ExpandProperty CAPID
    )

    return @(
        $trackCapIds + $dutyCapIds + $commanderCapIds |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { "$_".Trim() } |
            Sort-Object -Unique
    )
}

function Get-UnitCodeFromCompanyName {
    param (
        [string]$companyName
    )

    if ([string]::IsNullOrWhiteSpace($companyName)) {
        return $null
    }

    if ($companyName -match '(?i)\bCO-(\d+)\b') {
        return $matches[1]
    }

    if ($companyName -match '\b(\d+)\b') {
        return $matches[1]
    }

    return $null
}

function Get-UserCapId {
    param (
        [object]$User
    )

    if ($null -eq $User) {
        return $null
    }

    if (-not [string]::IsNullOrWhiteSpace($User.employeeId)) {
        return "$($User.employeeId)".Trim()
    }

    return $null
}

function Test-UserCapIdInList {
    param (
        [object]$User,
        [array]$CapIds
    )

    $userCapId = Get-UserCapId -User $User
    $normalizedCapIds = @(
        $CapIds |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { "$_".Trim() }
    )

    return -not [string]::IsNullOrWhiteSpace($userCapId) -and $normalizedCapIds -contains $userCapId
}

function Test-WingMatches {
    param (
        [string]$RecordWing,
        [string]$WingDesignator
    )

    if ([string]::IsNullOrWhiteSpace($WingDesignator)) {
        return $true
    }

    if ([string]::IsNullOrWhiteSpace($RecordWing)) {
        return $false
    }

    $recordWingNormalized = $RecordWing.Trim().ToUpperInvariant()
    $wingDesignatorNormalized = $WingDesignator.Trim().ToUpperInvariant()

    return $recordWingNormalized -eq $wingDesignatorNormalized -or $wingDesignatorNormalized.StartsWith($recordWingNormalized)
}

function Get-CommanderCapIdsForUnit {
    param (
        [array]$commanders = @(),
        [string]$unit,
        [string]$wingDesignator = $null
    )

    return @(
        $commanders |
            Where-Object {
                -not [string]::IsNullOrWhiteSpace($_.CAPID) -and
                "$($_.Unit)".Trim() -eq "$unit".Trim() -and
                (Test-WingMatches -RecordWing $_.Wing -WingDesignator $wingDesignator)
            } |
            Select-Object -ExpandProperty CAPID |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { "$_".Trim() } |
            Sort-Object -Unique
    )
}

function Get-RecruitingMembersForUnit {
    param (
        [array]$allUsers,
        [string]$unit,
        [array]$recruitingCAPIDs,
        [array]$commanders = @(),
        [string]$wingDesignator = $null
    )

    $unitCommanderCapIds = @(Get-CommanderCapIdsForUnit -commanders $commanders -unit $unit -wingDesignator $wingDesignator)
    if ($unitCommanderCapIds.Count -gt 0) {
        Write-DLSpecTrackLog "Commanders.txt CAPIDs for unit $unit : $($unitCommanderCapIds -join ',')"
    } else {
        Write-DLSpecTrackLog "No Commanders.txt CAPIDs found for unit $unit"
    }

    foreach ($commanderCapId in $unitCommanderCapIds) {
        $commanderUsers = @($allUsers | Where-Object { (Get-UserCapId -User $_) -eq $commanderCapId })
        if ($commanderUsers.Count -eq 0) {
            Write-DLSpecTrackLog "Commanders.txt CAPID $commanderCapId for unit $unit did not match any Entra user employeeId"
            continue
        }

        foreach ($commanderUser in $commanderUsers) {
            $commanderUserUnit = Get-UnitCodeFromCompanyName -companyName $commanderUser.companyName
            $hasMail = $null -ne $commanderUser.mail
            Write-DLSpecTrackLog "Commanders.txt CAPID $commanderCapId matched Entra user '$($commanderUser.displayName)' employeeId='$($commanderUser.employeeId)' companyName='$($commanderUser.companyName)' extractedUnit='$commanderUserUnit' hasMail=$hasMail"
        }
    }

    return @(
        $allUsers |
            Where-Object {
                $userUnit = Get-UnitCodeFromCompanyName -companyName $_.companyName
                $isUnitMatch = $userUnit -eq $unit
                $isUnitCommander = Test-UserCapIdInList -User $_ -CapIds $unitCommanderCapIds

                (
                    ($isUnitMatch -and (Test-UserCapIdInList -User $_ -CapIds $recruitingCAPIDs)) -or
                    $isUnitCommander -or
                    ($isUnitMatch -and $_.department -like '*EX*')
                ) -and
                $_.mail -ne $null
            } |
            Sort-Object -Property id -Unique
    )
}

function Get-RecruitingGroupMailNickname {
    param (
        [string]$UnitCode
    )

    return "co-$($UnitCode.ToLowerInvariant())-recruiting"
}

function Get-RecruitingGroupAddress {
    param (
        [string]$UnitCode
    )

    return "$(Get-RecruitingGroupMailNickname -UnitCode $UnitCode)@cowg.cap.gov"
}

function Test-RecruitingGroupAddressNeedsRename {
    param (
        [string]$CurrentAddress,
        [string]$DesiredAddress
    )

    if ([string]::IsNullOrWhiteSpace($CurrentAddress) -or [string]::IsNullOrWhiteSpace($DesiredAddress)) {
        return $false
    }

    return $CurrentAddress.Trim().ToLowerInvariant() -ne $DesiredAddress.Trim().ToLowerInvariant()
}

function Test-RecruitingGroupRequiresExternalSenderUpdate {
    param (
        [string]$GroupName,
        [object]$Group
    )

    if ($GroupName -notmatch '(?i)^CO-\d+ Recruiting$') {
        return $false
    }

    if ($null -eq $Group) {
        return $true
    }

    $property = $Group.PSObject.Properties['RequireSenderAuthenticationEnabled']
    if ($null -eq $property) {
        return $true
    }

    return $property.Value -ne $false
}

function Get-DistributionGroupTypeForName {
    param (
        [string]$GroupName
    )

    if ($GroupName -match '(?i)^CO-\d+ Recruiting$') {
        return "Security"
    }

    return "Distribution"
}

function Test-RecruitingGroupRequiresSecurityMigration {
    param (
        [string]$GroupName,
        [object]$Group
    )

    if ($GroupName -notmatch '(?i)^CO-\d+ Recruiting$' -or $null -eq $Group) {
        return $false
    }

    $recipientTypeDetails = $Group.PSObject.Properties['RecipientTypeDetails']
    if ($null -ne $recipientTypeDetails) {
        return "$($recipientTypeDetails.Value)" -ne "MailUniversalSecurityGroup"
    }

    $groupType = $Group.PSObject.Properties['GroupType']
    if ($null -ne $groupType) {
        return "$($groupType.Value)" -notmatch '(?i)Security'
    }

    return $true
}

function Get-DistributionGroupMemberUpdateOptions {
    param (
        [string]$GroupName
    )

    $options = @{
        ErrorAction = "Stop"
    }

    if ($GroupName -match '(?i)^CO-\d+ Recruiting$') {
        $options.BypassSecurityGroupManagerCheck = $true
    }

    return $options
}

function Get-DistributionGroupRemovalOptions {
    param (
        [string]$GroupName
    )

    $options = @{
        Confirm = $false
        ErrorAction = "Stop"
    }

    if ($GroupName -match '(?i)^CO-\d+ Recruiting$') {
        $options.BypassSecurityGroupManagerCheck = $true
    }

    return $options
}

function Get-GraphGroupMembersUri {
    param (
        [string]$GroupId
    )

    return "https://graph.microsoft.com/v1.0/groups/$GroupId/members?`$select=id"
}

function Get-GraphGroupLookupMaxAttempts {
    param (
        [string]$GroupName
    )

    if ($GroupName -match '(?i)^CO-\d+ Recruiting$') {
        return 6
    }

    return 1
}

function Get-GraphGroupLookupDelaySeconds {
    return 5
}

function Write-DLSpecTrackLog {
    param (
        [string]$Message
    )

    if (Get-Command Write-Log -ErrorAction SilentlyContinue) {
        Write-Log $Message
    }
}

# This function compares two arrays and returns the user IDs that are in both, only in the first array, and only in the second array.
function Compare-Arrays {
    param (
        [array]$Array1, # Full user objects from the filtered list
        [array]$Array2  # IDs of current group members
    )

    $Array1 = @(
        $Array1 |
            Where-Object { $null -ne $_ -and -not [string]::IsNullOrWhiteSpace($_.id) } |
            Sort-Object -Property id -Unique
    )

    $Array2 = @(
        $Array2 |
            Where-Object { $null -ne $_ -and -not [string]::IsNullOrWhiteSpace([string]$_) } |
            ForEach-Object { ([string]$_).Trim().ToLowerInvariant() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
    )

    $inBoth = @($Array1 | Where-Object { $Array2 -contains $_.id.Trim().ToLowerInvariant() })
    Write-DLSpecTrackLog "InBoth count: $($inBoth.Count)"

    $Add = @($Array1 | Where-Object { $Array2 -notcontains $_.id.Trim().ToLowerInvariant() })
    Write-DLSpecTrackLog "Add count: $($Add.Count)"

    $Array1Hash = @{}
    foreach ($user in $Array1) {
        $Array1Hash[$user.id.Trim().ToLowerInvariant()] = $user
    }

    if ($Array1Hash.Count -gt 0) {
        $Remove = @($Array2 | Where-Object { -not $Array1Hash.ContainsKey($_) })
        Write-DLSpecTrackLog "Remove count: $($Remove.Count)"
    } else {
        $Remove = @()
        Write-DLSpecTrackLog "No users to remove, Array1Hash is empty."
    }

    return [PSCustomObject]@{
        InBoth = $inBoth
        Add    = $Add
        Remove = $Remove
    }
}
