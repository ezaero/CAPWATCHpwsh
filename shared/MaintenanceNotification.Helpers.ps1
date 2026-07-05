function ConvertTo-GraphEmailRecipients {
    param(
        [array]$Recipients
    )

    $seen = @{}
    $graphRecipients = @()

    foreach ($recipient in $Recipients) {
        $address = if ($recipient) { "$recipient".Trim() } else { "" }
        if ([string]::IsNullOrWhiteSpace($address)) {
            continue
        }

        $key = $address.ToLowerInvariant()
        if ($seen.ContainsKey($key)) {
            continue
        }

        $seen[$key] = $true
        $graphRecipients += @{
            emailAddress = @{
                address = $address
            }
        }
    }

    return $graphRecipients
}

function Get-MaintenanceNotificationUnit {
    param(
        [object]$Record
    )

    $rawUnit = $null
    if ($Record -and $Record.PSObject.Properties.Name -contains "Unit") {
        $rawUnit = $Record.Unit
    }
    if ([string]::IsNullOrWhiteSpace([string]$rawUnit) -and $Record -and $Record.PSObject.Properties.Name -contains "companyName") {
        $rawUnit = $Record.companyName
    }

    $unit = if ($rawUnit) { "$rawUnit".Trim() } else { "" }
    if ([string]::IsNullOrWhiteSpace($unit)) {
        return "Unknown"
    }

    if ($unit -match '^(?i)CO-(\d{1,3})$') {
        return "CO-{0:D3}" -f [int]$matches[1]
    }

    if ($unit -match '^\d{1,3}$') {
        return "CO-{0:D3}" -f [int]$unit
    }

    return $unit
}

function Split-MaintenanceDeletedDisplayName {
    param(
        [string]$DisplayName
    )

    $name = if ($DisplayName) { $DisplayName.Trim() } else { "" }
    if ([string]::IsNullOrWhiteSpace($name)) {
        return [PSCustomObject]@{
            NameFirst = ""
            NameLast = ""
        }
    }

    $parts = $name -split '\s+'
    return [PSCustomObject]@{
        NameFirst = $parts[0]
        NameLast = if ($parts.Count -gt 1) { $parts[1..($parts.Count - 1)] -join " " } else { "" }
    }
}

function Get-MaintenanceMemberUnitByCapid {
    param(
        [string]$Capid,
        [array]$MemberRows
    )

    if ([string]::IsNullOrWhiteSpace($Capid) -or -not $MemberRows) {
        return "Unknown"
    }

    $baseCapid = "$Capid" -replace 'P$', ''
    $member = $MemberRows | Where-Object { "$($_.CAPID)" -eq $baseCapid } | Select-Object -First 1
    return Get-MaintenanceNotificationUnit -Record $member
}

function ConvertFrom-MaintenanceDeletionLog {
    param(
        [array]$LogLines,
        [array]$MemberRows = @()
    )

    $deletedMembers = @()
    $seen = @{}

    foreach ($line in $LogLines) {
        $record = $null

        if ($line -match 'Deleted (member|parent) account: (.+?), ([^()]+) \(([^)]+)\) with CAPID: ([^.]+)\.') {
            $kind = $matches[1]
            $displayName = $matches[2].Trim()
            $grade = $matches[3].Trim()
            $email = $matches[4].Trim()
            $capid = $matches[5].Trim()
            $nameParts = Split-MaintenanceDeletedDisplayName -DisplayName $displayName

            $record = [PSCustomObject]@{
                NameFirst = $nameParts.NameFirst
                NameLast = $nameParts.NameLast
                Grade = $grade
                CAPID = $capid
                Email = $email
                Unit = Get-MaintenanceMemberUnitByCapid -Capid $capid -MemberRows $MemberRows
                Type = if ($kind -eq "parent") { "Parent" } else { "Member" }
            }
        } elseif ($line -match 'Deleted O365 account: (.+?) \(([^)]+)\), CAPID: ([^,]+), Unit: (.+)$') {
            $displayAndGrade = $matches[1].Trim()
            $email = $matches[2].Trim()
            $capid = $matches[3].Trim()
            $unit = $matches[4].Trim()
            $displayName = $displayAndGrade
            $grade = "Unknown"

            if ($displayAndGrade -match '^(.+?),\s*(.+)$') {
                $displayName = $matches[1].Trim()
                $grade = $matches[2].Trim()
            }

            $nameParts = Split-MaintenanceDeletedDisplayName -DisplayName $displayName

            $record = [PSCustomObject]@{
                NameFirst = $nameParts.NameFirst
                NameLast = $nameParts.NameLast
                Grade = $grade
                CAPID = $capid
                Email = $email
                Unit = Get-MaintenanceNotificationUnit -Record ([PSCustomObject]@{ Unit = $unit })
                Type = "Inactive"
            }
        }

        if ($record) {
            $key = "{0}|{1}|{2}" -f $record.CAPID, $record.Email.ToLowerInvariant(), $record.Type
            if (-not $seen.ContainsKey($key)) {
                $seen[$key] = $true
                $deletedMembers += $record
            }
        }
    }

    return $deletedMembers
}
