$ErrorActionPreference = "Stop"

. "$PSScriptRoot/../shared/MaintenanceNotification.Helpers.ps1"

function Assert-Equal {
    param($Actual, $Expected, [string]$Message)
    if ($Actual -ne $Expected) {
        throw "$Message Expected '$Expected' but got '$Actual'."
    }
}

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) {
        throw $Message
    }
}

$graphRecipients = ConvertTo-GraphEmailRecipients -Recipients @(
    "mike.schulte@cowg.cap.gov",
    " commander@example.org ",
    "",
    $null,
    "mike.schulte@cowg.cap.gov"
)

Assert-Equal $graphRecipients.Count 2 "Graph recipients should skip blanks and de-duplicate addresses."
Assert-Equal $graphRecipients[0].emailAddress.address "mike.schulte@cowg.cap.gov" "First Graph recipient should keep the first valid address."
Assert-Equal $graphRecipients[1].emailAddress.address "commander@example.org" "Graph recipient addresses should be trimmed."

$json = @{ message = @{ toRecipients = $graphRecipients } } | ConvertTo-Json -Depth 6
Assert-True ($json -match '"emailAddress"') "Graph recipients should serialize as emailAddress objects."
Assert-True ($json -notmatch '\["mike\.schulte@cowg\.cap\.gov"') "Graph recipients should not serialize as plain strings."

Assert-Equal (Get-MaintenanceNotificationUnit -Record ([PSCustomObject]@{ Unit = "186" })) "CO-186" "CAPWATCH Unit should be normalized to CO-###."
Assert-Equal (Get-MaintenanceNotificationUnit -Record ([PSCustomObject]@{ Unit = "CO-186" })) "CO-186" "Existing companyName-style units should be preserved."
Assert-Equal (Get-MaintenanceNotificationUnit -Record ([PSCustomObject]@{ companyName = "CO-186" })) "CO-186" "companyName fallback should be supported."
Assert-Equal (Get-MaintenanceNotificationUnit -Record ([PSCustomObject]@{ Unit = ""; companyName = "" })) "Unknown" "Blank unit data should not produce CO-."

$memberRows = @(
    [PSCustomObject]@{ CAPID = "138597"; Unit = "022" },
    [PSCustomObject]@{ CAPID = "639580"; Unit = "030" }
)

$deletedFromLog = ConvertFrom-MaintenanceDeletionLog -LogLines @(
    "2026-07-03 22:00:45 - Deleted member account: John Burke, SM (John.Burke@cowg.cap.gov) with CAPID: 138597.",
    "2026-07-03 22:00:53 - Deleted parent account: Amber Strachan, SM PARENT (yvonnestrachan10@gmail.com) with CAPID: 639580P.",
    "2026-07-03 22:01:51 - Deleted O365 account: AACS Testing (aacstesting@cowg.cap.gov), CAPID: none, Unit: CO-159",
    "2026-07-03 22:02:50 - Failed to send expired members notification email for unit CO-001"
) -MemberRows $memberRows

Assert-Equal $deletedFromLog.Count 3 "Deletion log replay should return only deleted account lines."
Assert-Equal $deletedFromLog[0].NameFirst "John" "Member deletion should parse first name."
Assert-Equal $deletedFromLog[0].NameLast "Burke" "Member deletion should parse last name."
Assert-Equal $deletedFromLog[0].Grade "SM" "Member deletion should parse grade."
Assert-Equal $deletedFromLog[0].CAPID "138597" "Member deletion should parse CAPID."
Assert-Equal $deletedFromLog[0].Email "John.Burke@cowg.cap.gov" "Member deletion should parse email."
Assert-Equal $deletedFromLog[0].Unit "CO-022" "Member deletion should resolve unit from CAPWATCH rows."
Assert-Equal $deletedFromLog[0].Type "Member" "Member deletion should identify member type."
Assert-Equal $deletedFromLog[1].Unit "CO-030" "Parent deletion should resolve unit using the base CAPID."
Assert-Equal $deletedFromLog[1].Type "Parent" "Parent deletion should identify parent type."
Assert-Equal $deletedFromLog[2].Unit "CO-159" "O365 deletion should use the unit already present in the log."
Assert-Equal $deletedFromLog[2].Type "Inactive" "O365 deletion should identify inactive type."

Write-Host "Maintenance notification helper tests passed."
