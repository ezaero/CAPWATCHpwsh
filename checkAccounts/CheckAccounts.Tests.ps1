$ErrorActionPreference = "Stop"

. "$PSScriptRoot\CheckAccounts.Helpers.ps1"

function Assert-Equal {
    param(
        [object]$Actual,
        [object]$Expected,
        [string]$Message
    )

    if ($Actual -ne $Expected) {
        throw "$Message Expected '$Expected', got '$Actual'."
    }
}

$contact = [PSCustomObject]@{
    CAPID = "211568"
    Unit = "164"
}

$commanders = @(
    [PSCustomObject]@{
        CAPID = "211568"
        Wing = "CO"
        Unit = "186"
    },
    [PSCustomObject]@{
        CAPID = "211568"
        Wing = "AL"
        Unit = "024"
    }
)

$desiredCompanyName = Get-DesiredCompanyName -Contact $contact -Commanders $commanders -WingDesignator "CO"
Assert-Equal $desiredCompanyName "CO-186" "Commanders.txt unit should override stale member unit for commander companyName."

$nonCommanderContact = [PSCustomObject]@{
    CAPID = "999999"
    Unit = "164"
}

$fallbackCompanyName = Get-DesiredCompanyName -Contact $nonCommanderContact -Commanders $commanders -WingDesignator "CO"
Assert-Equal $fallbackCompanyName "CO-164" "Non-commanders should keep their member unit companyName."

$memberWithoutContactFlag = [PSCustomObject]@{
    CAPID = "211568"
    Unit = "186"
    DoNotContact = $null
    Type = "SENIOR"
    MbrStatus = "ACTIVE"
    Email = $null
}

Assert-Equal (Test-IncludeMemberForAccountSync -Member $memberWithoutContactFlag) $true "Active members should still be checked for account attributes when DoNotContact is null."

$doNotContactMember = [PSCustomObject]@{
    CAPID = "211568"
    Unit = "186"
    DoNotContact = "True"
    Type = "SENIOR"
    MbrStatus = "ACTIVE"
    Email = "Charles.Sellers@cowg.cap.gov"
}

Assert-Equal (Test-IncludeMemberForAccountSync -Member $doNotContactMember) $true "DoNotContact=True members should still be checked for account attributes."

$seniorMember = [PSCustomObject]@{
    CAPID = "767094"
    Type = "SENIOR"
    Email = "derrick841@msn.com"
}

$parentAccountByCapid = [PSCustomObject]@{
    employeeId = "763701P"
    employeeType = $null
    jobTitle = "C/Amn"
}

Assert-Equal (Test-ParentCommunicationAccount -User $parentAccountByCapid) $true "Accounts with parent CAPID suffix should be treated as parent communication accounts."
Assert-Equal (Test-ShouldReplaceParentAccountForSeniorMember -Member $seniorMember -ExistingUser $parentAccountByCapid) $true "Senior member should be allowed to replace parent account with conflicting email."

$parentAccountByJobTitle = [PSCustomObject]@{
    employeeId = $null
    employeeType = $null
    jobTitle = "C/TSgt PARENT"
}

Assert-Equal (Test-ParentCommunicationAccount -User $parentAccountByJobTitle) $true "Accounts with parent job titles should be treated as parent communication accounts."

$nonParentAccount = [PSCustomObject]@{
    employeeId = "767094"
    employeeType = "SENIOR"
    jobTitle = "SM"
}

Assert-Equal (Test-ShouldReplaceParentAccountForSeniorMember -Member $seniorMember -ExistingUser $nonParentAccount) $false "Senior member should not replace non-parent accounts."

$parentAccountByProxyAddress = [PSCustomObject]@{
    employeeId = "763701P"
    employeeType = "PARENT"
    jobTitle = "C/Amn PARENT"
    mail = $null
    proxyAddresses = @("smtp:derrick841@msn.com")
}

Assert-Equal (Test-DirectoryUserEmailConflict -User $parentAccountByProxyAddress -Email "Derrick841@msn.com") $true "Email conflict detection should match proxyAddresses case-insensitively."
Assert-Equal (Test-ShouldReplaceParentAccountForSeniorMember -Member $seniorMember -ExistingUser $parentAccountByProxyAddress) $true "Senior member should be allowed to replace recognizable parent accounts found by proxy address."

$nonParentAccountByProxyAddress = [PSCustomObject]@{
    employeeId = "767094"
    employeeType = "SENIOR"
    jobTitle = "SM"
    mail = $null
    proxyAddresses = @("SMTP:derrick841@msn.com")
}

Assert-Equal (Test-DirectoryUserEmailConflict -User $nonParentAccountByProxyAddress -Email "derrick841@msn.com") $true "Email conflict detection should identify non-parent proxy conflicts."
Assert-Equal (Test-ShouldReplaceParentAccountForSeniorMember -Member $seniorMember -ExistingUser $nonParentAccountByProxyAddress) $false "Senior member should not replace non-parent accounts even when the email conflicts."

$cadetMember = [PSCustomObject]@{
    CAPID = "767094"
    Type = "CADET"
    Email = "derrick841@msn.com"
}

Assert-Equal (Test-ShouldReplaceParentAccountForSeniorMember -Member $cadetMember -ExistingUser $parentAccountByCapid) $false "Non-senior members should not replace parent accounts automatically."

$permanentDeleteUri = Get-DirectoryDeletedItemPermanentDeleteUri -ObjectId "9005a3b9-ff7e-446c-9bb8-754e26d38747"
Assert-Equal $permanentDeleteUri "https://graph.microsoft.com/v1.0/directory/deletedItems/9005a3b9-ff7e-446c-9bb8-754e26d38747" "Permanent delete should target the deletedItems endpoint by object id."

$missingPermanentDeleteUri = Get-DirectoryDeletedItemPermanentDeleteUri -ObjectId ""
Assert-Equal $missingPermanentDeleteUri $null "Permanent delete URI should be null when object id is missing."

$blankExchangeRecipient = [PSCustomObject]@{}
Assert-Equal (Test-ExchangeRecipientConflictObject -Recipient $blankExchangeRecipient) $false "Blank Exchange output should not be treated as a real recipient conflict."

$realExchangeRecipient = [PSCustomObject]@{
    Identity = "7c47efa6-6126-4ae4-8301-961ed427626f"
    RecipientTypeDetails = "GuestMailUser"
}
Assert-Equal (Test-ExchangeRecipientConflictObject -Recipient $realExchangeRecipient) $true "Exchange objects with identity and type should be treated as real recipient conflicts."

$manualGuest = ConvertTo-ManualGuestRecord -Row ([PSCustomObject]@{
    Email = "kaitlin.reinhart@utwg.cap.gov"
    DisplayName = "Kaitlin Reinhart"
    Grade = "Winter Encampment Staff"
    Unit = "WINTER-ENCAMPMENT"
    Teams = "Winter Encampment 2026;CO-200 Example"
    Sponsor = "Maj Example Sponsor"
    Enabled = "TRUE"
})
Assert-Equal $manualGuest.CAPID "MANUAL-1BA42EFA87483DA9" "Manual guest employeeId should be stable from normalized email."
Assert-Equal $manualGuest.NameFirst "Kaitlin" "Manual guest first name should be derived from DisplayName when needed."
Assert-Equal $manualGuest.NameLast "Reinhart" "Manual guest last name should be derived from DisplayName when needed."
Assert-Equal $manualGuest.Type "MANUAL" "Manual guest type should default to MANUAL."
Assert-Equal $manualGuest.Sponsor "Maj Example Sponsor" "Manual guest sponsor should be retained for audit logging."
Assert-Equal ((Split-ManualGuestTeams -Teams $manualGuest.Teams).Count) 2 "Manual guest Teams should split on semicolon delimiters."

$disabledManualGuest = ConvertTo-ManualGuestRecord -Row ([PSCustomObject]@{
    Email = "disabled@example.org"
    Enabled = "FALSE"
})
Assert-Equal $disabledManualGuest $null "Disabled manual guest rows should be ignored."

# Duplicate cleanup must retain a live internal member and only target a parent
# guest identified by both CAPID and exact contact email (including secondary).
$cleanupSenior = [PSCustomObject]@{ CAPID = '100001'; Type = 'SENIOR'; Unit = '191'; Email = 'adult@cowg.cap.gov' }
$cleanupParent = [PSCustomObject]@{ CAPID = '200001P'; Type = 'PARENT'; Email = 'personal@example.org' }
$cleanupContacts = @([PSCustomObject]@{ CAPID = '100001'; Type = 'EMAIL'; Priority = 'SECONDARY'; Contact = ' Personal@Example.org ' })
$cleanupMemberUser = [PSCustomObject]@{ id = 'keep'; employeeId = '100001'; userType = 'Member'; mail = 'adult@cowg.cap.gov' }
$cleanupGuest = [PSCustomObject]@{ id = 'delete'; employeeId = '200001P'; userType = 'Guest'; mail = 'personal@example.org' }
function Get-TestCleanupPlan {
    param([array]$Users = @($cleanupMemberUser, $cleanupGuest), [array]$Info = @($cleanupSenior, $cleanupParent), [array]$Contacts = $cleanupContacts)
    Get-DuplicateParentGuestPlan -MemberInfo $Info -Contacts $Contacts -AllUsers $Users
}
$cleanupPlans = @(Get-TestCleanupPlan)
Assert-Equal $cleanupPlans.Count 1 'Exact secondary email should link differently addressed parent and member accounts.'
Assert-Equal $cleanupPlans[0].GuestUser.id 'delete' 'Only the parent guest should be deleted.'
Assert-Equal $cleanupPlans[0].RetainedUser.id 'keep' 'Internal member must be retained.'
$afterDeletion = @(Get-TestCleanupPlan -Users @($cleanupMemberUser))
Assert-Equal $afterDeletion.Count 1 'Parent must remain suppressed after deletion to prevent recreation.'
Assert-Equal $afterDeletion[0].GuestUser $null 'An already absent guest must not be deleted again.'
Assert-Equal (@(Get-TestCleanupPlan -Users @($cleanupGuest)).Count) 0 'No cleanup without retained member account.'
$guestSenior = [PSCustomObject]@{ id = 'keep'; employeeId = '100001'; userType = 'Guest' }
Assert-Equal (@(Get-TestCleanupPlan -Users @($guestSenior, $cleanupGuest)).Count) 0 'A senior guest alone must not authorize parent deletion.'
$internalParent = [PSCustomObject]@{ id = 'delete'; employeeId = '200001P'; userType = 'Member'; mail = 'personal@example.org' }
Assert-Equal (@(Get-TestCleanupPlan -Users @($cleanupMemberUser, $internalParent)).Count) 0 'Never delete an internal parent account.'
$changedGuest = [PSCustomObject]@{ id = 'delete'; employeeId = '200001P'; userType = 'Guest'; mail = 'someoneelse@example.org' }
Assert-Equal (@(Get-TestCleanupPlan -Users @($cleanupMemberUser, $changedGuest)).Count) 0 'Mismatched guest email requires review.'
Assert-Equal (@(Get-TestCleanupPlan -Contacts @()).Count) 0 'Names or phone numbers alone cannot authorize deletion.'
Assert-Equal (@(Get-TestCleanupPlan -Users @($cleanupMemberUser, $cleanupGuest, $changedGuest)).Count) 0 'Ambiguous parent CAPIDs must be skipped.'
$secondSenior = [PSCustomObject]@{ CAPID = '100002'; Type = 'SENIOR'; Unit = '191'; Email = 'spouse@cowg.cap.gov' }
$sharedContacts = $cleanupContacts + [PSCustomObject]@{ CAPID = '100002'; Type = 'EMAIL'; Contact = 'personal@example.org' }
Assert-Equal (@(Get-TestCleanupPlan -Info @($cleanupSenior, $secondSenior, $cleanupParent) -Contacts $sharedContacts).Count) 0 'Email shared by two senior records must not authorize deletion.'

Write-Host "CheckAccounts.Tests.ps1 passed"
