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

Write-Host "CheckAccounts.Tests.ps1 passed"
