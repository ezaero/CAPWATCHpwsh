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

function Get-CommanderUnitForCapId {
    param (
        [array]$Commanders = @(),
        [string]$CapId,
        [string]$WingDesignator
    )

    $commander = $Commanders |
        Where-Object {
            "$($_.CAPID)".Trim() -eq "$CapId".Trim() -and
            -not [string]::IsNullOrWhiteSpace($_.Unit) -and
            (Test-WingMatches -RecordWing $_.Wing -WingDesignator $WingDesignator)
        } |
        Select-Object -First 1

    if ($commander) {
        return "$($commander.Unit)".Trim()
    }

    return $null
}

function Get-DesiredCompanyName {
    param (
        [object]$Contact,
        [array]$Commanders = @(),
        [string]$WingDesignator = "CO"
    )

    $unit = Get-CommanderUnitForCapId -Commanders $Commanders -CapId $Contact.CAPID -WingDesignator $WingDesignator
    if ([string]::IsNullOrWhiteSpace($unit)) {
        $unit = "$($Contact.Unit)".Trim()
    }

    if ([string]::IsNullOrWhiteSpace($WingDesignator)) {
        $WingDesignator = "CO"
    }

    return "$($WingDesignator.Trim().ToUpperInvariant())-$unit"
}

function Test-IncludeMemberForAccountSync {
    param (
        [object]$Member
    )

    return $Member.Unit -ne "999" -and
        $Member.Unit -ne "000" -and
        $Member.Type -ne "AEM" -and
        $Member.Type -ne "PATRON" -and
        $Member.MbrStatus -ne "EXPIRED" -and
        -not ($Member.Email -and $Member.Email -match '(?i)@coloradomilitaryacademy\.org$')
}
