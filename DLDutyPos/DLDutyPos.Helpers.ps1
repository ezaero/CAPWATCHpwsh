. "$PSScriptRoot\..\DLSeniorsCadets\DLSeniorsCadets.Helpers.ps1"

$script:DLDutyPosGraphRoot = "https://graph.microsoft.com/v1.0"

function Get-DLDutyPosConfigPath {
    param (
        [string]$ConfiguredPath = $env:DLDUTYPOS_CONFIG_PATH
    )

    if (-not [string]::IsNullOrWhiteSpace($ConfiguredPath)) {
        return $ConfiguredPath
    }

    $yamlPath = Join-Path -Path $PSScriptRoot -ChildPath "distribution-lists.yaml"
    if (Test-Path -Path $yamlPath) {
        return $yamlPath
    }

    return Join-Path -Path $PSScriptRoot -ChildPath "distribution-lists.json"
}

function ConvertFrom-DLDutyPosYamlScalar {
    param (
        [AllowNull()]
        [string]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    $trimmed = $Value.Trim()
    if ($trimmed -eq "[]") {
        return @()
    }

    if ($trimmed -match '^(?i:true)$') {
        return $true
    }

    if ($trimmed -match '^(?i:false)$') {
        return $false
    }

    if ($trimmed -match '^(?i:null)$') {
        return $null
    }

    if (
        ($trimmed.StartsWith('"') -and $trimmed.EndsWith('"')) -or
        ($trimmed.StartsWith("'") -and $trimmed.EndsWith("'"))
    ) {
        return $trimmed.Substring(1, $trimmed.Length - 2)
    }

    return $trimmed
}

function ConvertFrom-DLDutyPosYaml {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Yaml
    )

    $definitions = @()
    $currentDefinition = $null
    $currentListProperty = $null

    foreach ($rawLine in ($Yaml -split "`r?`n")) {
        if ([string]::IsNullOrWhiteSpace($rawLine)) {
            continue
        }

        $lineWithoutComment = $rawLine -replace '\s+#.*$', ''
        if ([string]::IsNullOrWhiteSpace($lineWithoutComment)) {
            continue
        }

        $indent = ($lineWithoutComment.Length - $lineWithoutComment.TrimStart().Length)
        $line = $lineWithoutComment.Trim()

        if ($line -eq "distributionLists:") {
            continue
        }

        if ($line.StartsWith("- ")) {
            $item = $line.Substring(2).Trim()
            if ($indent -le 2 -or $null -eq $currentDefinition) {
                if ($null -ne $currentDefinition) {
                    $definitions += [PSCustomObject]$currentDefinition
                }

                $currentDefinition = [ordered]@{}
                $currentListProperty = $null

                if (-not [string]::IsNullOrWhiteSpace($item)) {
                    if ($item -notmatch '^([A-Za-z][A-Za-z0-9]*):\s*(.*)$') {
                        throw "Unsupported DLDutyPos YAML list item '$line'."
                    }

                    $currentDefinition[$matches[1]] = ConvertFrom-DLDutyPosYamlScalar -Value $matches[2]
                }

                continue
            }

            if ([string]::IsNullOrWhiteSpace($currentListProperty)) {
                throw "DLDutyPos YAML list item '$line' is not under a list property."
            }

            $currentDefinition[$currentListProperty] += @(ConvertFrom-DLDutyPosYamlScalar -Value $item)
            continue
        }

        if ($null -eq $currentDefinition) {
            continue
        }

        if ($line -notmatch '^([A-Za-z][A-Za-z0-9]*):\s*(.*)$') {
            throw "Unsupported DLDutyPos YAML line '$line'."
        }

        $key = $matches[1]
        $value = $matches[2]
        if ([string]::IsNullOrWhiteSpace($value)) {
            $currentDefinition[$key] = @()
            $currentListProperty = $key
        } else {
            $currentDefinition[$key] = ConvertFrom-DLDutyPosYamlScalar -Value $value
            $currentListProperty = $null
        }
    }

    if ($null -ne $currentDefinition) {
        $definitions += [PSCustomObject]$currentDefinition
    }

    return @($definitions)
}

function Import-DLDutyPosConfig {
    param (
        [string]$Path = (Get-DLDutyPosConfigPath)
    )

    if (-not (Test-Path -Path $Path)) {
        return @()
    }

    $rawConfig = Get-Content -Path $Path -Raw
    $extension = [System.IO.Path]::GetExtension($Path)
    if ($extension -match '^(?i:\.ya?ml)$') {
        return @(ConvertFrom-DLDutyPosYaml -Yaml $rawConfig)
    }

    $config = $rawConfig | ConvertFrom-Json
    if ($config -is [array]) {
        return @($config)
    }

    if ($config.distributionLists) {
        return @($config.distributionLists)
    }

    return @($config)
}

function ConvertTo-DLDutyPosAlias {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $alias = $Value -replace "[^A-Za-z0-9-]", ""
    if ($alias.Length -gt 64) {
        return $alias.Substring(0, 64)
    }

    return $alias
}

function ConvertTo-DLDutyPosGraphFilterLiteral {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    return $Value.Replace("'", "''")
}

function Get-DLDutyPosGraphGroupByDisplayName {
    param (
        [Parameter(Mandatory = $true)]
        [string]$DisplayName
    )

    $escaped = ConvertTo-DLDutyPosGraphFilterLiteral -Value $DisplayName
    $uri = "$script:DLDutyPosGraphRoot/groups?`$select=id,displayName,mailNickname&`$filter=displayName eq '$escaped'"
    $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
    $groups = @($response.value)
    if ($groups.Count -gt 1) {
        throw "Multiple Entra groups found with displayName '$DisplayName'."
    }

    if ($groups.Count -eq 0) {
        return $null
    }

    return $groups[0]
}

function Get-DLDutyPosGraphGroupMemberIds {
    param (
        [Parameter(Mandatory = $true)]
        [string]$GroupId
    )

    $memberIds = @()
    $uri = "$script:DLDutyPosGraphRoot/groups/$GroupId/members/microsoft.graph.user?`$select=id"
    while (-not [string]::IsNullOrWhiteSpace($uri)) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        $memberIds += @($response.value | ForEach-Object { $_.id } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $uri = $response.'@odata.nextLink'
    }

    return @($memberIds | Sort-Object -Unique)
}

function Get-DLDutyPosGraphGroupDisplayNamesByPrefix {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Prefixes
    )

    $displayNames = @()
    foreach ($prefix in @($Prefixes | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)) {
        $escaped = ConvertTo-DLDutyPosGraphFilterLiteral -Value $prefix
        $uri = "$script:DLDutyPosGraphRoot/groups?`$select=displayName&`$filter=startswith(displayName,'$escaped')"
        while (-not [string]::IsNullOrWhiteSpace($uri)) {
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            $displayNames += @($response.value | ForEach-Object { $_.displayName } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            $uri = $response.'@odata.nextLink'
        }
    }

    return @($displayNames | Sort-Object -Unique)
}

function Get-DLDutyPosSourceGroupMembers {
    param (
        [Parameter(Mandatory = $true)]
        [array]$GroupNames
    )

    $membersByGroup = @{}
    $missingGroups = @()

    foreach ($groupName in @($GroupNames | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        $group = Get-DLDutyPosGraphGroupByDisplayName -DisplayName $groupName
        if ($null -eq $group) {
            $missingGroups += $groupName
            continue
        }

        $membersByGroup[$groupName] = @(Get-DLDutyPosGraphGroupMemberIds -GroupId $group.id)
    }

    return [PSCustomObject]@{
        MembersByGroup = $membersByGroup
        MissingGroups = @($missingGroups | Sort-Object -Unique)
    }
}

function Get-DLDutyPosDefinitionBool {
    param (
        [object]$Definition,
        [string]$Name,
        [bool]$Default = $false
    )

    $property = $Definition.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) {
        return $Default
    }

    if ($property.Value -is [bool]) {
        return $property.Value
    }

    return "$($property.Value)" -match '^(?i:true|yes|1)$'
}

function New-DLDutyPosDistributionGroupPlan {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Definition,

        [Parameter(Mandatory = $true)]
        [hashtable]$SourceMembersByGroup,

        [Parameter(Mandatory = $true)]
        [hashtable]$MailByUserId
    )

    $includeIds = @()
    foreach ($groupName in @($Definition.includeGroups)) {
        if ($SourceMembersByGroup.ContainsKey($groupName)) {
            $includeIds += @($SourceMembersByGroup[$groupName])
        }
    }

    $excludeIds = @()
    foreach ($groupName in @($Definition.excludeGroups)) {
        if ($SourceMembersByGroup.ContainsKey($groupName)) {
            $excludeIds += @($SourceMembersByGroup[$groupName])
        }
    }

    $desiredIds = @(
        $includeIds |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique |
            Where-Object { $excludeIds -notcontains $_ }
    )

    $members = @(
        foreach ($userId in $desiredIds) {
            if ($MailByUserId.ContainsKey($userId) -and -not [string]::IsNullOrWhiteSpace($MailByUserId[$userId])) {
                $MailByUserId[$userId]
            }
        }
    ) | Sort-Object -Unique

    $displayName = $Definition.displayName
    $email = $Definition.email
    if ([string]::IsNullOrWhiteSpace($email)) {
        throw "Distribution list definition '$displayName' is missing an email value."
    }

    $alias = $Definition.alias
    if ([string]::IsNullOrWhiteSpace($alias)) {
        $alias = ($email -split "@")[0]
    }
    $alias = ConvertTo-DLDutyPosAlias -Value $alias
    $ensureGroup = Get-DLDutyPosDefinitionBool -Definition $Definition -Name "ensureGroup" -Default $false

    return New-DLGroupUpdatePlan `
        -Identity $displayName `
        -Members $members `
        -Category "DLDutyPos" `
        -EnsureGroup $ensureGroup `
        -Name $displayName `
        -Alias $alias `
        -EmailAddress $email `
        -UpdateIdentity $email
}

function New-DLDutyPosDistributionGroupPlans {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Definitions,

        [Parameter(Mandatory = $true)]
        [array]$AllUsers,

        [hashtable]$SourceMembersByGroup,

        [array]$MissingGroups = @()
    )

    $mailByUserId = @{}
    foreach ($user in $AllUsers) {
        if ($null -eq $user -or [string]::IsNullOrWhiteSpace($user.id)) {
            continue
        }

        $mail = $user.mail
        if ([string]::IsNullOrWhiteSpace($mail)) {
            $mail = $user.userPrincipalName
        }

        if (-not [string]::IsNullOrWhiteSpace($mail)) {
            $mailByUserId["$($user.id)".Trim()] = "$mail".Trim()
        }
    }

    $plans = @()
    foreach ($definition in $Definitions) {
        $includeGroups = @($definition.includeGroups | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $excludeGroups = @($definition.excludeGroups | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $referencedGroups = @($includeGroups + $excludeGroups | Sort-Object -Unique)
        $definitionMissingGroups = @($referencedGroups | Where-Object { $MissingGroups -contains $_ })
        $allowMissingSourceGroups = $definition.allowMissingSourceGroups -eq $true

        if ($definitionMissingGroups.Count -gt 0 -and -not $allowMissingSourceGroups) {
            if (Get-Command Write-Log -ErrorAction SilentlyContinue) {
                Write-Log "Skipping DLDutyPos list '$($definition.displayName)' because source groups are missing: $($definitionMissingGroups -join ', ')"
            }
            continue
        }

        $plans += New-DLDutyPosDistributionGroupPlan -Definition $definition -SourceMembersByGroup $SourceMembersByGroup -MailByUserId $mailByUserId
    }

    return @($plans | Sort-Object Identity)
}

function Get-DLDutyPosSourceGroupNames {
    param (
        [array]$Definitions
    )

    return @(
        foreach ($definition in $Definitions) {
            @($definition.includeGroups)
            @($definition.excludeGroups)
        }
    ) |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique
}

function Get-DLDutyPosPatternPrefix {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Pattern
    )

    $match = [regex]::Match($Pattern, "[\*\?\[]")
    if (-not $match.Success) {
        return $Pattern
    }

    return $Pattern.Substring(0, $match.Index)
}

function Get-DLDutyPosSourceGroupPatternPrefixes {
    param (
        [array]$Definitions
    )

    return @(
        foreach ($definition in $Definitions) {
            @($definition.includeGroupPatterns)
            @($definition.excludeGroupPatterns)
        }
    ) |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object { Get-DLDutyPosPatternPrefix -Pattern $_ } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique
}

function Resolve-DLDutyPosDefinitionGroupPatterns {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Definitions,

        [Parameter(Mandatory = $true)]
        [array]$CandidateGroupNames
    )

    $candidates = @($CandidateGroupNames | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
    foreach ($definition in $Definitions) {
        $copy = [ordered]@{}
        foreach ($property in $definition.PSObject.Properties) {
            $copy[$property.Name] = $property.Value
        }

        $includeGroups = @($definition.includeGroups | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        foreach ($pattern in @($definition.includeGroupPatterns | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
            $includeGroups += @($candidates | Where-Object { $_ -like $pattern })
        }

        $excludeGroups = @($definition.excludeGroups | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        foreach ($pattern in @($definition.excludeGroupPatterns | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
            $excludeGroups += @($candidates | Where-Object { $_ -like $pattern })
        }

        $copy["includeGroups"] = @($includeGroups | Sort-Object -Unique)
        $copy["excludeGroups"] = @($excludeGroups | Sort-Object -Unique)
        [PSCustomObject]$copy
    }
}

