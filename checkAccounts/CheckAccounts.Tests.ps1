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

Write-Host "CheckAccounts.Tests.ps1 passed"
