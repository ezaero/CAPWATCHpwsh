function Test-ValidManualGuestEmailAddress {
    param (
        [string]$EmailAddress
    )

    return -not [string]::IsNullOrWhiteSpace($EmailAddress) -and
        $EmailAddress.Trim() -match '^[\w\.\-]+@([\w\-]+\.)+[\w\-]{2,}$'
}

function New-ManualGuestEmployeeId {
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress
    )

    $normalizedEmail = $EmailAddress.Trim().ToLowerInvariant()
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($normalizedEmail)
    $hash = [System.Security.Cryptography.SHA256]::Create().ComputeHash($bytes)
    $hex = -join ($hash | ForEach-Object { $_.ToString("x2") })

    return "MANUAL-$($hex.Substring(0, 16).ToUpperInvariant())"
}

function Split-ManualGuestTeams {
    param (
        [string]$Teams
    )

    if ([string]::IsNullOrWhiteSpace($Teams)) {
        return @()
    }

    return @(
        $Teams -split '[;|]' |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
    )
}

function ConvertTo-ManualGuestRecord {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Row
    )

    $enabled = if ($Row.PSObject.Properties.Name -contains "Enabled") { "$($Row.Enabled)".Trim() } else { "" }
    if ($enabled -match '^(?i:false|no|0|disabled)$') {
        return $null
    }

    $email = if ($Row.Email) { "$($Row.Email)".Trim() } else { "" }
    if (-not (Test-ValidManualGuestEmailAddress -EmailAddress $email)) {
        return $null
    }

    $displayName = if ($Row.PSObject.Properties.Name -contains "DisplayName") { "$($Row.DisplayName)".Trim() } else { "" }
    $nameFirst = if ($Row.PSObject.Properties.Name -contains "NameFirst") { "$($Row.NameFirst)".Trim() } else { "" }
    $nameLast = if ($Row.PSObject.Properties.Name -contains "NameLast") { "$($Row.NameLast)".Trim() } else { "" }

    if (([string]::IsNullOrWhiteSpace($nameFirst) -or [string]::IsNullOrWhiteSpace($nameLast)) -and -not [string]::IsNullOrWhiteSpace($displayName)) {
        $nameParts = @($displayName -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ([string]::IsNullOrWhiteSpace($nameFirst) -and $nameParts.Count -gt 0) {
            $nameFirst = $nameParts[0]
        }
        if ([string]::IsNullOrWhiteSpace($nameLast) -and $nameParts.Count -gt 1) {
            $nameLast = $nameParts[1..($nameParts.Count - 1)] -join ' '
        }
    }

    if ([string]::IsNullOrWhiteSpace($nameFirst)) {
        $nameFirst = ($email -split '@')[0]
    }
    if ([string]::IsNullOrWhiteSpace($nameLast)) {
        $nameLast = "Guest"
    }

    $employeeId = if ($Row.PSObject.Properties.Name -contains "EmployeeId") { "$($Row.EmployeeId)".Trim() } else { "" }
    if ([string]::IsNullOrWhiteSpace($employeeId)) {
        $employeeId = New-ManualGuestEmployeeId -EmailAddress $email
    }

    $unit = if ($Row.PSObject.Properties.Name -contains "Unit") { "$($Row.Unit)".Trim() } else { "" }
    if ([string]::IsNullOrWhiteSpace($unit)) {
        $unit = "MANUAL"
    }

    $grade = if ($Row.PSObject.Properties.Name -contains "Grade") { "$($Row.Grade)".Trim() } else { "" }
    if ([string]::IsNullOrWhiteSpace($grade)) {
        $grade = "Manual Guest"
    }

    $type = if ($Row.PSObject.Properties.Name -contains "Type") { "$($Row.Type)".Trim() } else { "" }
    if ([string]::IsNullOrWhiteSpace($type)) {
        $type = "MANUAL"
    }

    $teams = if ($Row.PSObject.Properties.Name -contains "Teams") { "$($Row.Teams)" } else { "" }
    $sponsor = if ($Row.PSObject.Properties.Name -contains "Sponsor") { "$($Row.Sponsor)".Trim() } else { "" }

    return [PSCustomObject]@{
        CAPID        = $employeeId
        EmployeeId   = $employeeId
        NameLast     = $nameLast
        NameFirst    = $nameFirst
        Unit         = $unit
        Grade        = $grade
        Type         = $type
        Email        = $email
        DoNotContact = $false
        MobilePhone  = if ($Row.PSObject.Properties.Name -contains "MobilePhone") { "$($Row.MobilePhone)".Trim() } else { $null }
        DOB          = if ($Row.PSObject.Properties.Name -contains "DOB") { "$($Row.DOB)".Trim() } else { $null }
        Joined       = if ($Row.PSObject.Properties.Name -contains "Joined") { "$($Row.Joined)".Trim() } else { $null }
        MbrStatus    = "ACTIVE"
        Source       = "ManualGuests"
        Teams        = $teams
        Sponsor      = $sponsor
    }
}

function Get-ManualGuestsCsvPath {
    $projectRoot = Split-Path -Parent $PSScriptRoot
    return Join-Path -Path $projectRoot -ChildPath "ManualGuests.csv"
}

function Import-ManualGuestRows {
    param (
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        $Path = Get-ManualGuestsCsvPath
    }

    if (-not (Test-Path -Path $Path)) {
        return @()
    }

    $records = @()
    foreach ($row in (Import-Csv -Path $Path -ErrorAction Stop)) {
        $record = ConvertTo-ManualGuestRecord -Row $row
        if ($record) {
            $records += $record
        }
    }

    return @($records | Sort-Object Email -Unique)
}
