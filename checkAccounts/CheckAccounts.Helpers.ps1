if (-not (Get-Command Import-ManualGuestRows -ErrorAction SilentlyContinue)) {
    . "$PSScriptRoot\..\shared\ManualGuests.Helpers.ps1"
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

function Get-DuplicateParentGuestPlan {
    param (
        [array]$MemberInfo,
        [array]$Contacts,
        [array]$AllUsers
    )

    # Only exact CAPWATCH email evidence links identities. Names/phones can belong
    # to spouses. Require an existing internal member account before suppressing
    # a parent identity, including on subsequent runs after the guest is gone.
    $seniors = @{}
    foreach ($member in $MemberInfo) {
        if ($member.Type -eq 'SENIOR' -and (Test-IncludeMemberForAccountSync -Member $member)) {
            $seniors["$($member.CAPID)"] = $member
        }
    }
    $emailOwners = @{}
    foreach ($contact in $Contacts) {
        if ($contact.Type -ne 'EMAIL' -or -not $seniors.ContainsKey("$($contact.CAPID)")) { continue }
        $email = "$($contact.Contact)".Trim().ToLowerInvariant()
        if ($email -notmatch '^[^\s@]+@[^\s@]+\.[^\s@]+$') { continue }
        if (-not $emailOwners.ContainsKey($email)) { $emailOwners[$email] = @{} }
        $emailOwners[$email]["$($contact.CAPID)"] = $true
    }
    foreach ($parent in $MemberInfo) {
        if ($parent.Type -ne 'PARENT' -or $parent.CAPID -notmatch '^\d+P$') { continue }
        $email = "$($parent.Email)".Trim().ToLowerInvariant()
        if (-not $emailOwners.ContainsKey($email) -or $emailOwners[$email].Count -ne 1) { continue }
        $seniorCapid = @($emailOwners[$email].Keys)[0]
        $retained = @($AllUsers | Where-Object {
            $_.employeeId -eq $seniorCapid -and $_.userType -eq 'Member' -and $_.id
        })
        if ($retained.Count -ne 1) { continue }
        $parentUsers = @($AllUsers | Where-Object { $_.employeeId -eq $parent.CAPID })
        # An unexpected type, ambiguous CAPID, or changed email needs review.
        if ($parentUsers.Count -gt 1) { continue }
        $guest = $null
        if ($parentUsers.Count -eq 1) {
            $guest = $parentUsers[0]
            if ($guest.userType -ne 'Guest' -or -not $guest.id -or
                $guest.id -eq $retained[0].id -or
                -not (Test-DirectoryUserEmailConflict -User $guest -Email $email)) { continue }
        }
        [PSCustomObject]@{
            ParentCapid = $parent.CAPID
            ParentEmail = $email
            SeniorCapid = $seniorCapid
            RetainedUser = $retained[0]
            GuestUser = $guest
        }
    }
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

function Test-DirectoryUserEmailConflict {
    param (
        [object]$User,
        [string]$Email,
        [string]$UserPrincipalName = $null
    )

    if (-not $User -or [string]::IsNullOrWhiteSpace($Email)) {
        return $false
    }

    $emailNormalized = $Email.Trim().ToLowerInvariant()
    $userPrincipalNameNormalized = if (-not [string]::IsNullOrWhiteSpace($UserPrincipalName)) {
        $UserPrincipalName.Trim().ToLowerInvariant()
    } else {
        $null
    }

    if ($User.mail -and $User.mail.ToString().Trim().ToLowerInvariant() -eq $emailNormalized) {
        return $true
    }

    if ($userPrincipalNameNormalized -and
        $User.userPrincipalName -and
        $User.userPrincipalName.ToString().Trim().ToLowerInvariant() -eq $userPrincipalNameNormalized) {
        return $true
    }

    if ($User.proxyAddresses) {
        foreach ($proxyAddress in @($User.proxyAddresses)) {
            if ($null -eq $proxyAddress) {
                continue
            }

            $proxyAddressText = $proxyAddress.ToString().Trim()
            if ($proxyAddressText -match '(?i)^smtp:(.+)$' -and $Matches[1].Trim().ToLowerInvariant() -eq $emailNormalized) {
                return $true
            }
        }
    }

    return $false
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
