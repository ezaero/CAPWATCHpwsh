function Get-DLGroupNumber {
    param (
        [Parameter(Mandatory = $true)]
        [object]$GroupOrganization
    )

    if ($GroupOrganization.Name -match '(?i)\bGROUP\s+(\d+)\b') {
        return $matches[1]
    }

    if ($GroupOrganization.Unit -match '^0*(\d+)$') {
        return $matches[1]
    }

    return $null
}

function Write-DLHelperLog {
    param (
        [string]$Message
    )

    if (Get-Command Write-Log -ErrorAction SilentlyContinue) {
        Write-Log $Message
    }
}

function Get-RegionalDistributionGroups {
    param (
        [Parameter(Mandatory = $true)]
        [array]$OrganizationRows,

        [string]$EmailDomain = "cowg.cap.gov"
    )

    $activeWingOrganizations = @(
        $OrganizationRows |
            Where-Object {
                $_.Wing -eq $env:WING_DESIGNATOR -and
                $_.Status -eq "ACTIVE"
            }
    )

    $activeGroups = @(
        $activeWingOrganizations |
            Where-Object { $_.Scope -eq "GROUP" -or $_.Type -eq "GROUP" } |
            Sort-Object Name
    )

    foreach ($group in $activeGroups) {
        $groupNumber = Get-DLGroupNumber -GroupOrganization $group
        if ([string]::IsNullOrWhiteSpace($groupNumber)) {
            Write-DLHelperLog "Skipping CAPWATCH group organization '$($group.Name)' because no group number could be resolved."
            continue
        }

        $alias = "group$groupNumber"
        $units = @(
            $activeWingOrganizations |
                Where-Object {
                    $_.Scope -eq "UNIT" -and
                    $_.NextLevel -eq $group.ORGID -and
                    $_.Unit -notin @("000", "001", "999")
                } |
                Select-Object -ExpandProperty Unit |
                Sort-Object -Unique
        )

        if ($units.Count -eq 0) {
            Write-DLHelperLog "Skipping $($group.Name) because CAPWATCH has no active child units for ORGID $($group.ORGID)."
            continue
        }

        [PSCustomObject]@{
            Name = "$($env:WING_DESIGNATOR) Group $groupNumber"
            Alias = $alias
            EmailAddress = "$alias@$EmailDomain"
            Units = $units
            GroupOrgId = $group.ORGID
            GroupUnit = $group.Unit
            GroupName = $group.Name
        }
    }
}

function Get-RegionalDistributionGroupMembers {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Group,

        [Parameter(Mandatory = $true)]
        [array]$AllUsers
    )

    $companyNames = $Group.Units | ForEach-Object { "$($env:WING_DESIGNATOR)-$_" }
    return @(
        $AllUsers |
            Where-Object { $companyNames -contains $_.companyName } |
            Select-Object -ExpandProperty mail |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
    )
}

function ConvertTo-DLMemberEmails {
    param (
        [array]$Members
    )

    return @(
        $Members |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
    )
}

function New-DLGroupUpdatePlan {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Identity,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [array]$Members,

        [string]$Category = "General",

        [bool]$EnsureGroup = $false,

        [string]$Name = "",

        [string]$Alias = "",

        [string]$EmailAddress = ""
    )

    return [PSCustomObject]@{
        Identity     = $Identity
        Members      = @(ConvertTo-DLMemberEmails -Members $Members)
        Category     = $Category
        EnsureGroup  = $EnsureGroup
        Name         = $Name
        Alias        = $Alias
        EmailAddress = $EmailAddress
    }
}

function Get-DLSquadronGroupPlans {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateSet("SENIOR", "CADET", "ALL")]
        [string]$MemberType,

        [Parameter(Mandatory = $true)]
        [array]$UnitList,

        [Parameter(Mandatory = $true)]
        [array]$AllUsers
    )

    foreach ($unit in $UnitList) {
        $unitDesignator = "$($env:WING_DESIGNATOR)-$($unit.Unit)"

        if ($MemberType -eq "ALL") {
            $groupName = "$($env:WING_DESIGNATOR)-$($unit.Unit) $($unit.Name)"
            $members = @($AllUsers | Where-Object { $_.companyName -eq $unitDesignator } | Select-Object -ExpandProperty mail)
        } else {
            $memberName = ($MemberType.Substring(0, 1).ToUpper()) + ($MemberType.Substring(1).ToLower()) + "s"
            $groupName = "$($env:WING_DESIGNATOR)-$($unit.Unit) $memberName"
            $members = @($AllUsers | Where-Object { $_.companyName -eq $unitDesignator -and $_.employeeType -eq $MemberType } | Select-Object -ExpandProperty mail)

            if ($MemberType -eq "CADET") {
                $members += @($AllUsers | Where-Object { $_.companyName -eq $unitDesignator -and $_.employeeType -eq "PARENT" } | Select-Object -ExpandProperty mail)
                $members += @($AllUsers | Where-Object { $_.companyName -eq $unitDesignator -and ($_.department -like "*EX*" -or $_.department -like "*CP*") } | Select-Object -ExpandProperty mail)
            }
        }

        New-DLGroupUpdatePlan -Identity $groupName -Members $members -Category "Squadron$MemberType"
    }
}

function Get-DLWingGroupPlan {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateSet("SENIOR", "CADET")]
        [string]$MemberType,

        [Parameter(Mandatory = $true)]
        [array]$AllUsers
    )

    $groupName = "$($env:WING_DESIGNATOR) Wing $MemberType`s"
    $members = @()

    if ($MemberType -eq "CADET") {
        $members += @($AllUsers | Where-Object { $_.employeeType -eq "CADET" } | Select-Object -ExpandProperty mail)
        $members += @($AllUsers | Where-Object { $_.employeeType -eq "PARENT" } | Select-Object -ExpandProperty mail)
        $members += @($AllUsers | Where-Object { $_.department -like "*CP*" } | Select-Object -ExpandProperty mail)
        $members += @($AllUsers | Where-Object { $_.department -like "*EX*" } | Select-Object -ExpandProperty mail)
    } else {
        $members += @($AllUsers | Where-Object { $_.employeeType -eq "SENIOR" } | Select-Object -ExpandProperty mail)
    }

    New-DLGroupUpdatePlan -Identity $groupName -Members $members -Category "Wing$MemberType"
}

function Get-DLRegionalGroupPlans {
    param (
        [Parameter(Mandatory = $true)]
        [array]$AllUsers,

        [Parameter(Mandatory = $true)]
        [array]$OrganizationRows
    )

    foreach ($group in (Get-RegionalDistributionGroups -OrganizationRows $OrganizationRows)) {
        $members = @(Get-RegionalDistributionGroupMembers -Group $group -AllUsers $AllUsers)
        New-DLGroupUpdatePlan `
            -Identity $group.EmailAddress `
            -Members $members `
            -Category "Regional" `
            -EnsureGroup $true `
            -Name $group.Name `
            -Alias $group.Alias `
            -EmailAddress $group.EmailAddress
    }
}

function New-DLGroupUpdatePlans {
    param (
        [Parameter(Mandatory = $true)]
        [array]$UnitList,

        [Parameter(Mandatory = $true)]
        [array]$AllUsers,

        [Parameter(Mandatory = $true)]
        [array]$OrganizationRows
    )

    return @(
        Get-DLSquadronGroupPlans -MemberType "SENIOR" -UnitList $UnitList -AllUsers $AllUsers
        Get-DLSquadronGroupPlans -MemberType "CADET" -UnitList $UnitList -AllUsers $AllUsers
        Get-DLSquadronGroupPlans -MemberType "ALL" -UnitList $UnitList -AllUsers $AllUsers
        Get-DLWingGroupPlan -MemberType "CADET" -AllUsers $AllUsers
        Get-DLWingGroupPlan -MemberType "SENIOR" -AllUsers $AllUsers
        Get-DLRegionalGroupPlans -AllUsers $AllUsers -OrganizationRows $OrganizationRows
    )
}

function ConvertTo-DLSafeFileName {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    return ($Value -replace '[^A-Za-z0-9._-]', '_').Trim("_")
}

function Save-DLGroupUpdateJob {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Plan,

        [Parameter(Mandatory = $true)]
        [string]$RunId,

        [string]$RootPath = (Join-Path $env:HOME "data/DLSeniorsCadets/jobs")
    )

    $runPath = Join-Path $RootPath $RunId
    if (-not (Test-Path $runPath)) {
        New-Item -ItemType Directory -Path $runPath -Force | Out-Null
    }

    $safeIdentity = ConvertTo-DLSafeFileName -Value $Plan.Identity
    $jobPath = Join-Path $runPath "$safeIdentity.json"

    $job = [PSCustomObject]@{
        RunId        = $RunId
        Identity     = $Plan.Identity
        Members      = @($Plan.Members)
        Category     = $Plan.Category
        EnsureGroup  = $Plan.EnsureGroup
        Name         = $Plan.Name
        Alias        = $Plan.Alias
        EmailAddress = $Plan.EmailAddress
        CreatedAtUtc = (Get-Date).ToUniversalTime().ToString("o")
    }

    $job | ConvertTo-Json -Depth 6 | Set-Content -Path $jobPath -Encoding utf8

    return [PSCustomObject]@{
        RunId       = $RunId
        Identity    = $Plan.Identity
        Category    = $Plan.Category
        JobPath     = $jobPath
        MemberCount = @($Plan.Members).Count
    }
}

function Ensure-DLDistributionGroup {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Job
    )

    if (-not $Job.EnsureGroup) {
        return
    }

    try {
        $null = Get-DistributionGroup -Identity $Job.EmailAddress -ErrorAction Stop
        Write-Log "Distribution group already exists: $($Job.Name) ($($Job.EmailAddress))"
    } catch {
        $null = New-DistributionGroup `
            -Name $Job.Name `
            -DisplayName $Job.Name `
            -Alias $Job.Alias `
            -PrimarySmtpAddress $Job.EmailAddress `
            -Type Distribution `
            -ErrorAction Stop
        Write-Log "Distribution group created: $($Job.Name) ($($Job.EmailAddress))"
    }
}

function Get-DLCurrentDistributionGroupMembers {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )

    return @(
        Get-DistributionGroupMember -Identity $Identity -ResultSize Unlimited -ErrorAction Stop |
            ForEach-Object {
                if ($_.PrimarySmtpAddress) {
                    $_.PrimarySmtpAddress.ToString()
                } elseif ($_.WindowsEmailAddress) {
                    $_.WindowsEmailAddress.ToString()
                } elseif ($_.ExternalEmailAddress) {
                    $_.ExternalEmailAddress.ToString()
                }
            } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
    )
}

function Test-DLExchangeObjectNotFound {
    param (
        [Parameter(Mandatory = $true)]
        [object]$ErrorObject
    )

    $message = $ErrorObject.ToString()
    if ($ErrorObject.Exception -and $ErrorObject.Exception.Message) {
        $message = $ErrorObject.Exception.Message
    }

    return ($message -match "could(n't| not) be found" -or $message -match "cannot be found" -or $message -match "was not found")
}

function Compare-DLMemberEmails {
    param (
        [array]$DesiredMembers,
        [array]$CurrentMembers
    )

    $desired = @(ConvertTo-DLMemberEmails -Members $DesiredMembers)
    $current = @(ConvertTo-DLMemberEmails -Members $CurrentMembers)

    return [PSCustomObject]@{
        Desired = $desired
        Current = $current
        Add     = @($desired | Where-Object { $current -notcontains $_ })
        Remove  = @($current | Where-Object { $desired -notcontains $_ })
    }
}

function Invoke-DLGroupUpdateJob {
    param (
        [Parameter(Mandatory = $true)]
        [string]$JobPath
    )

    if (-not (Test-Path $JobPath)) {
        throw "DL group update job file was not found: $JobPath"
    }

    $job = Get-Content -Path $JobPath -Raw | ConvertFrom-Json
    $members = @(ConvertTo-DLMemberEmails -Members $job.Members)

    Ensure-DLDistributionGroup -Job $job

    try {
        $currentMembers = @(Get-DLCurrentDistributionGroupMembers -Identity $job.Identity)
    } catch {
        if (-not $job.EnsureGroup -and (Test-DLExchangeObjectNotFound -ErrorObject $_)) {
            Write-Log "Distribution group '$($job.Identity)' does not exist or could not be resolved. Skipping $($job.Category) job with $($members.Count) desired members."
            return
        }

        throw
    }
    $comparison = Compare-DLMemberEmails -DesiredMembers $members -CurrentMembers $currentMembers

    if ($comparison.Add.Count -eq 0 -and $comparison.Remove.Count -eq 0) {
        Write-Log "Distribution group '$($job.Identity)' already has $($members.Count) desired members. Skipping update."
        return
    }

    if (($comparison.Add.Count + $comparison.Remove.Count) -le 25) {
        foreach ($member in $comparison.Add) {
            Add-DistributionGroupMember -Identity $job.Identity -Member $member -ErrorAction Stop
        }

        foreach ($member in $comparison.Remove) {
            Remove-DistributionGroupMember -Identity $job.Identity -Member $member -Confirm:$false -ErrorAction Stop
        }

        Write-Log "Incrementally updated distribution group '$($job.Identity)': added $($comparison.Add.Count), removed $($comparison.Remove.Count), desired total $($members.Count)."
        return
    }

    Update-DistributionGroupMember -Identity $job.Identity -Members $members -Confirm:$false -ErrorAction Stop
    Write-Log "Replaced distribution group '$($job.Identity)' membership with $($members.Count) desired members."
}
