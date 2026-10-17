$ErrorActionPreference = "Stop"

. "$PSScriptRoot\DLDutyPos.Helpers.ps1"

function Assert-Equal {
    param (
        [object]$Actual,
        [object]$Expected,
        [string]$Message
    )

    if ($Actual -ne $Expected) {
        throw "$Message Expected '$Expected', got '$Actual'."
    }
}

function Assert-ArrayEqual {
    param (
        [array]$Actual,
        [array]$Expected,
        [string]$Message
    )

    $actualText = @($Actual) -join ","
    $expectedText = @($Expected) -join ","
    if ($actualText -ne $expectedText) {
        throw "$Message Expected '$expectedText', got '$actualText'."
    }
}

$definitions = @(
    [PSCustomObject]@{
        displayName = "CO Wing Staff"
        email = "staff@cowg.cap.gov"
        includeGroups = @("CAP-Duty-WING-CO-EX-Commander")
        includeGroupPatterns = @("CAP-Duty-WING-CO-IT-*")
        excludeGroups = @("CAP-MemberStatus-INACTIVE")
    },
    [PSCustomObject]@{
        displayName = "Skipped Missing"
        email = "skipped@cowg.cap.gov"
        includeGroups = @("CAP-Duty-WING-CO-DOES-NOT-EXIST")
        excludeGroups = @()
    }
)

$sourceMembersByGroup = @{
    "CAP-Duty-WING-CO-EX-Commander" = @("user-1", "user-2")
    "CAP-Duty-WING-CO-IT-Director-of-IT" = @("user-2", "user-3")
    "CAP-Duty-WING-CO-IT-Assistant-Director-of-IT" = @("user-4")
    "CAP-MemberStatus-INACTIVE" = @("user-3")
}

$allUsers = @(
    [PSCustomObject]@{ id = "user-1"; mail = "commander@example.test"; userPrincipalName = "commander-upn@example.test" },
    [PSCustomObject]@{ id = "user-2"; mail = "it@example.test"; userPrincipalName = "it-upn@example.test" },
    [PSCustomObject]@{ id = "user-3"; mail = "inactive@example.test"; userPrincipalName = "inactive-upn@example.test" },
    [PSCustomObject]@{ id = "user-4"; mail = ""; userPrincipalName = "fallback@example.test" }
)

$definitions = @(Resolve-DLDutyPosDefinitionGroupPatterns `
    -Definitions $definitions `
    -CandidateGroupNames @(
        "CAP-Duty-WING-CO-IT-Director-of-IT",
        "CAP-Duty-WING-CO-IT-Assistant-Director-of-IT",
        "CAP-Duty-WING-CO-PA-Public-Affairs-Officer"
    ))

$plans = @(New-DLDutyPosDistributionGroupPlans `
    -Definitions $definitions `
    -AllUsers $allUsers `
    -SourceMembersByGroup $sourceMembersByGroup `
    -MissingGroups @("CAP-Duty-WING-CO-DOES-NOT-EXIST"))

Assert-Equal $plans.Count 1 "Missing source groups should skip their DL definition by default."
Assert-Equal $plans[0].Identity "CO Wing Staff" "Plan should preserve the configured display name."
Assert-Equal $plans[0].EmailAddress "staff@cowg.cap.gov" "Plan should preserve the configured email."
Assert-Equal $plans[0].EnsureGroup $false "DLDutyPos should reuse existing distribution lists unless ensureGroup is configured."
Assert-ArrayEqual $plans[0].Members @("commander@example.test", "fallback@example.test", "it@example.test") "Plan should union include groups and patterns, subtract exclusions, and resolve mail."

$sourceNames = @(Get-DLDutyPosSourceGroupNames -Definitions $definitions)
Assert-ArrayEqual $sourceNames @("CAP-Duty-WING-CO-DOES-NOT-EXIST", "CAP-Duty-WING-CO-EX-Commander", "CAP-Duty-WING-CO-IT-Assistant-Director-of-IT", "CAP-Duty-WING-CO-IT-Director-of-IT", "CAP-MemberStatus-INACTIVE") "Source group names should include includes and excludes exactly once."

$patternPrefixes = @(Get-DLDutyPosSourceGroupPatternPrefixes -Definitions $definitions)
Assert-ArrayEqual $patternPrefixes @("CAP-Duty-WING-CO-IT-") "Pattern prefixes should stop before wildcard tokens."

$yamlPath = Join-Path ([System.IO.Path]::GetTempPath()) "dl-dutypos-test-$([guid]::NewGuid()).yaml"
try {
    @'
distributionLists:
  - displayName: "YAML Staff"
    email: "yaml-staff@example.test"
    ensureGroup: false
    allowMissingSourceGroups: true
    includeGroups:
      - "CAP-Duty-WING-CO-EX-Commander"
    includeGroupPatterns:
      - "CAP-Duty-WING-CO-IT-*"
    excludeGroups: []
'@ | Set-Content -Path $yamlPath

    $yamlDefinitions = @(Import-DLDutyPosConfig -Path $yamlPath)
    Assert-Equal $yamlDefinitions.Count 1 "YAML config should import one distribution list definition."
    Assert-Equal $yamlDefinitions[0].displayName "YAML Staff" "YAML config should preserve scalar string values."
    Assert-Equal $yamlDefinitions[0].ensureGroup $false "YAML config should parse boolean false values."
    Assert-Equal $yamlDefinitions[0].allowMissingSourceGroups $true "YAML config should parse boolean true values."
    Assert-ArrayEqual $yamlDefinitions[0].includeGroups @("CAP-Duty-WING-CO-EX-Commander") "YAML config should parse includeGroups arrays."
    Assert-ArrayEqual $yamlDefinitions[0].includeGroupPatterns @("CAP-Duty-WING-CO-IT-*") "YAML config should parse includeGroupPatterns arrays."
    Assert-ArrayEqual $yamlDefinitions[0].excludeGroups @() "YAML config should parse empty arrays."
} finally {
    if (Test-Path -Path $yamlPath) {
        Remove-Item -Path $yamlPath -Force
    }
}

$createPlan = New-DLDutyPosDistributionGroupPlan `
    -Definition ([PSCustomObject]@{
        displayName = "Create Me"
        email = "create-me@example.test"
        includeGroups = @("CAP-Duty-WING-CO-EX-Commander")
        excludeGroups = @()
        ensureGroup = $true
    }) `
    -SourceMembersByGroup $sourceMembersByGroup `
    -MailByUserId @{
        "user-1" = "commander@example.test"
        "user-2" = "it@example.test"
    }
Assert-Equal $createPlan.EnsureGroup $true "Definitions should be able to opt in to creating missing distribution lists."

Write-Host "DLDutyPos.Tests.ps1 passed"

