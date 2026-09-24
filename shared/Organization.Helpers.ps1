function Get-CapwatchOrganizationPath {
    param (
        [string]$CapwatchDataDir = (Join-Path $env:HOME "data/CAPWatch")
    )

    return Join-Path -Path $CapwatchDataDir -ChildPath "Organization.txt"
}

function Import-CapwatchOrganizationRows {
    param (
        [string]$Path,
        [string]$CapwatchDataDir = (Join-Path $env:HOME "data/CAPWatch")
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        $Path = Get-CapwatchOrganizationPath -CapwatchDataDir $CapwatchDataDir
    }

    if (-not (Test-Path -Path $Path)) {
        throw "CAPWATCH Organization.txt was not found at $Path"
    }

    return @(Import-Csv -Path $Path -ErrorAction Stop)
}

function Test-CapwatchAdministrativeUnitCode {
    param (
        [string]$Unit
    )

    return $Unit -in @("000", "001", "999")
}

function Get-CapwatchActiveWingOrganizations {
    param (
        [Parameter(Mandatory = $true)]
        [array]$OrganizationRows,

        [string]$WingDesignator = $env:WING_DESIGNATOR
    )

    if ([string]::IsNullOrWhiteSpace($WingDesignator)) {
        $WingDesignator = "CO"
    }

    return @(
        $OrganizationRows |
            Where-Object {
                $_.Wing -eq $WingDesignator -and
                $_.Status -eq "ACTIVE"
            } |
            Sort-Object Scope, Unit, Name
    )
}

function Get-CapwatchActiveWingUnits {
    param (
        [Parameter(Mandatory = $true)]
        [array]$OrganizationRows,

        [string]$WingDesignator = $env:WING_DESIGNATOR,

        [switch]$IncludeAdministrative
    )

    return @(
        Get-CapwatchActiveWingOrganizations -OrganizationRows $OrganizationRows -WingDesignator $WingDesignator |
            Where-Object {
                $_.Scope -eq "UNIT" -and
                ($IncludeAdministrative -or -not (Test-CapwatchAdministrativeUnitCode -Unit $_.Unit))
            } |
            Sort-Object Unit -Unique
    )
}

function Get-CapwatchGroupNumber {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Organization
    )

    if ($Organization.Name -match '(?i)\bGROUP\s+(\d+)\b') {
        return $matches[1]
    }

    return $null
}

function Get-CapwatchActiveWingGroups {
    param (
        [Parameter(Mandatory = $true)]
        [array]$OrganizationRows,

        [string]$WingDesignator = $env:WING_DESIGNATOR
    )

    return @(
        Get-CapwatchActiveWingOrganizations -OrganizationRows $OrganizationRows -WingDesignator $WingDesignator |
            Where-Object { $_.Scope -eq "GROUP" -or $_.Type -eq "GROUP" } |
            Sort-Object Name
    )
}

function Get-CapwatchActiveUnitCodesForGroupNumber {
    param (
        [Parameter(Mandatory = $true)]
        [array]$OrganizationRows,

        [Parameter(Mandatory = $true)]
        [string]$GroupNumber,

        [string]$WingDesignator = $env:WING_DESIGNATOR
    )

    $activeOrganizations = @(Get-CapwatchActiveWingOrganizations -OrganizationRows $OrganizationRows -WingDesignator $WingDesignator)
    $group = $activeOrganizations |
        Where-Object { ($_.Scope -eq "GROUP" -or $_.Type -eq "GROUP") -and (Get-CapwatchGroupNumber -Organization $_) -eq $GroupNumber } |
        Select-Object -First 1

    if (-not $group) {
        return @()
    }

    return @(
        $activeOrganizations |
            Where-Object {
                $_.Scope -eq "UNIT" -and
                $_.NextLevel -eq $group.ORGID -and
                -not (Test-CapwatchAdministrativeUnitCode -Unit $_.Unit)
            } |
            Sort-Object Unit |
            Select-Object -ExpandProperty Unit
    )
}

function ConvertTo-CapwatchCompanyName {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Unit,

        [string]$WingDesignator = $env:WING_DESIGNATOR
    )

    if ([string]::IsNullOrWhiteSpace($WingDesignator)) {
        $WingDesignator = "CO"
    }

    return "$($WingDesignator.Trim().ToUpperInvariant())-$($Unit.Trim())"
}

function ConvertTo-CapwatchCompanyNames {
    param (
        [array]$Units,
        [string]$WingDesignator = $env:WING_DESIGNATOR
    )

    return @(
        $Units |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { ConvertTo-CapwatchCompanyName -Unit $_ -WingDesignator $WingDesignator }
    )
}
