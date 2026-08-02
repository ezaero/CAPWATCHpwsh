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

function Test-ParentCommunicationAccount {
    param (
        [object]$User
    )

    if (-not $User) {
        return $false
    }

    $employeeId = if ($null -ne $User.employeeId) { "$($User.employeeId)".Trim() } else { "" }
    $employeeType = if ($null -ne $User.employeeType) { "$($User.employeeType)".Trim() } else { "" }
    $jobTitle = if ($null -ne $User.jobTitle) { "$($User.jobTitle)".Trim() } else { "" }

    return $employeeId -match '(?i)P$' -or
        $employeeType -match '(?i)\bPARENT\b' -or
        $jobTitle -match '(?i)\bPARENT\b'
}

function Test-SeniorMember {
    param (
        [object]$Member
    )

    if (-not $Member -or $null -eq $Member.Type) {
        return $false
    }

    return "$($Member.Type)" -match '(?i)\bSENIOR\b'
}

function Test-ShouldReplaceParentAccountForSeniorMember {
    param (
        [object]$Member,
        [object]$ExistingUser
    )

    return (Test-SeniorMember -Member $Member) -and
        (Test-ParentCommunicationAccount -User $ExistingUser)
}

function Get-DirectoryDeletedItemPermanentDeleteUri {
    param (
        [string]$ObjectId
    )

    if ([string]::IsNullOrWhiteSpace($ObjectId)) {
        return $null
    }

    return "https://graph.microsoft.com/v1.0/directory/deletedItems/$($ObjectId.Trim())"
}

function Test-ExchangeRecipientConflictObject {
    param (
        [object]$Recipient
    )

    if (-not $Recipient) {
        return $false
    }

    $identity = if ($null -ne $Recipient.Identity) { "$($Recipient.Identity)".Trim() } else { "" }
    $name = if ($null -ne $Recipient.Name) { "$($Recipient.Name)".Trim() } else { "" }
    $recipientType = if ($null -ne $Recipient.RecipientType) { "$($Recipient.RecipientType)".Trim() } else { "" }
    $recipientTypeDetails = if ($null -ne $Recipient.RecipientTypeDetails) { "$($Recipient.RecipientTypeDetails)".Trim() } else { "" }

    return (-not [string]::IsNullOrWhiteSpace($identity) -or -not [string]::IsNullOrWhiteSpace($name)) -and
        (-not [string]::IsNullOrWhiteSpace($recipientType) -or -not [string]::IsNullOrWhiteSpace($recipientTypeDetails))
}
