param(
    [string]$Organization = $env:EXCHANGE_ORGANIZATION,
    [string]$ExchangeUserPrincipalName,
    [int]$BatchSize = 200,
    [int]$SleepSeconds = 15,
    [int]$MaxPasses = 0,
    [int]$MaxFailuresPerRecipient = 3,
    [switch]$ManagedIdentity,
    [switch]$UseDeviceCode,
    [switch]$SkipExchange,
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$modulePath = Join-Path $repoRoot "Modules"
if (Test-Path $modulePath) {
    $env:PSModulePath = "$modulePath$([IO.Path]::PathSeparator)$env:PSModulePath"
}

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
if (-not $SkipExchange) {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
}

function Write-BackfillLog {
    param([string]$Message)
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message"
}

function Connect-BackfillGraph {
    $context = Get-MgContext -ErrorAction SilentlyContinue
    if ($context) {
        Write-BackfillLog "Using existing Microsoft Graph connection for tenant $($context.TenantId)."
        return
    }

    if ($UseDeviceCode) {
        try {
            Import-Module Az.Accounts -ErrorAction Stop
            Connect-AzAccount -UseDeviceAuthentication -ErrorAction Stop | Out-Null
            $token = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token
            if ($token) {
                Connect-MgGraph -AccessToken $token -NoWelcome | Out-Null
                Write-BackfillLog "Connected to Microsoft Graph using a fresh Az device-code token."
                return
            }
        } catch {
            Write-BackfillLog "Could not connect to Microsoft Graph using Az device-code token: $($_.Exception.Message)"
        }
    }

    if (-not $UseDeviceCode) {
        try {
            Import-Module Az.Accounts -ErrorAction Stop
            $token = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token
            if ($token) {
                Connect-MgGraph -AccessToken $token -NoWelcome | Out-Null
                Write-BackfillLog "Connected to Microsoft Graph using the current Az account token."
                return
            }
        } catch {
            Write-BackfillLog "Could not connect to Microsoft Graph using Az token: $($_.Exception.Message)"
        }
    }

    if ($ManagedIdentity) {
        Connect-MgGraph -Identity -NoWelcome | Out-Null
        Write-BackfillLog "Connected to Microsoft Graph using managed identity."
        return
    }

    if ($UseDeviceCode) {
        Connect-MgGraph -Scopes "User.ReadWrite.All" -UseDeviceCode -NoWelcome | Out-Null
    } else {
        Connect-MgGraph -Scopes "User.ReadWrite.All" -NoWelcome | Out-Null
    }
    Write-BackfillLog "Connected to Microsoft Graph interactively."
}

function Connect-BackfillExchange {
    $connection = Get-ConnectionInformation -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($connection) {
        Write-BackfillLog "Using existing Exchange Online connection for $($connection.Organization)."
        return
    }

    if ($ManagedIdentity) {
        if ([string]::IsNullOrWhiteSpace($Organization)) {
            throw "Organization is required when using -ManagedIdentity."
        }
        Connect-ExchangeOnline -ManagedIdentity -Organization $Organization -ShowBanner:$false
        Write-BackfillLog "Connected to Exchange Online using managed identity for $Organization."
        return
    }

    if ($UseDeviceCode) {
        if (-not [string]::IsNullOrWhiteSpace($ExchangeUserPrincipalName)) {
            Connect-ExchangeOnline -UserPrincipalName $ExchangeUserPrincipalName -Device -ShowBanner:$false
        } else {
            Connect-ExchangeOnline -Device -ShowBanner:$false
        }
    } elseif (-not [string]::IsNullOrWhiteSpace($ExchangeUserPrincipalName)) {
        Connect-ExchangeOnline -UserPrincipalName $ExchangeUserPrincipalName -ShowBanner:$false
    } else {
        Connect-ExchangeOnline -ShowBanner:$false
    }
    Write-BackfillLog "Connected to Exchange Online."
}

function Get-GraphGuestUserMap {
    $guestUsers = @{}
    $guestUserObjects = @()
    $guestUserIds = @{}
    $uri = "https://graph.microsoft.com/beta/users?`$filter=userType eq 'Guest'&`$select=id,userPrincipalName,mail,employeeId,userType,showInAddressList&`$top=999"

    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
        foreach ($user in $response.value) {
            if ($user.id -and -not $guestUserIds.ContainsKey($user.id)) {
                $guestUserIds[$user.id] = $true
                $guestUserObjects += $user
            }

            foreach ($property in @("id", "userPrincipalName", "mail")) {
                if ($user.$property) {
                    $guestUsers["${property}:$($user.$property.ToString().ToLowerInvariant())"] = $user
                }
            }
        }
        $uri = $response.'@odata.nextLink'
    } while ($uri)

    return [PSCustomObject]@{
        Lookup = $guestUsers
        Users  = $guestUserObjects
    }
}

function Find-MatchedGuestUser {
    param(
        [object]$Recipient,
        [hashtable]$GuestUsers
    )

    if ($Recipient.ExternalDirectoryObjectId) {
        $matchedUser = $GuestUsers["id:$($Recipient.ExternalDirectoryObjectId.ToString().ToLowerInvariant())"]
        if ($matchedUser) { return $matchedUser }
    }

    if ($Recipient.UserPrincipalName) {
        $matchedUser = $GuestUsers["userPrincipalName:$($Recipient.UserPrincipalName.ToString().ToLowerInvariant())"]
        if ($matchedUser) { return $matchedUser }
    }

    if ($Recipient.PrimarySmtpAddress) {
        $matchedUser = $GuestUsers["mail:$($Recipient.PrimarySmtpAddress.ToString().ToLowerInvariant())"]
        if ($matchedUser) { return $matchedUser }
    }

    return $null
}

Connect-BackfillGraph
if (-not $SkipExchange) {
    Connect-BackfillExchange
} else {
    Write-BackfillLog "Skipping Exchange Online connection and HiddenFromAddressListsEnabled updates."
}

$failureCounts = @{}
$graphFailureCounts = @{}
$pass = 0

while ($true) {
    $pass++
    if ($MaxPasses -gt 0 -and $pass -gt $MaxPasses) {
        Write-BackfillLog "Stopping after MaxPasses=$MaxPasses."
        break
    }

    Write-BackfillLog "Starting pass $pass."
    $guestUserData = Get-GraphGuestUserMap
    $guestUsers = $guestUserData.Lookup
    $guestUserObjects = @($guestUserData.Users)
    Write-BackfillLog "Loaded $($guestUserObjects.Count) guest user(s) and $($guestUsers.Count) lookup key(s) from Microsoft Graph."

    $graphUpdates = @()
    foreach ($user in $guestUserObjects) {
        $capid = "$($user.employeeId)".Trim()
        $desiredShowInAddressList = -not ($capid -match '(?i)P$')
        $currentShowInAddressList = if ($null -eq $user.showInAddressList) { $null } else { $user.showInAddressList.ToString().ToLowerInvariant() }

        if ($null -ne $currentShowInAddressList -and $currentShowInAddressList -eq $desiredShowInAddressList.ToString().ToLowerInvariant()) {
            continue
        }

        if ($graphFailureCounts.ContainsKey($user.id) -and $graphFailureCounts[$user.id] -ge $MaxFailuresPerRecipient) {
            continue
        }

        $graphUpdates += [PSCustomObject]@{
            User                     = $user
            DesiredShowInAddressList = $desiredShowInAddressList
            Email                    = if ($user.mail) { $user.mail } else { $user.userPrincipalName }
            CAPID                    = $capid
        }
    }

    Write-BackfillLog "Found $($graphUpdates.Count) guest account(s) with incorrect Graph ShowInAddressList visibility."

    $graphBatch = if ($BatchSize -gt 0) { @($graphUpdates | Select-Object -First $BatchSize) } else { @($graphUpdates) }
    if ($graphBatch.Count -gt 0) {
        Write-BackfillLog "Updating $($graphBatch.Count) of $($graphUpdates.Count) Graph ShowInAddressList mismatch(es) this pass."
    }

    $graphProcessed = 0
    foreach ($item in $graphBatch) {
        $graphProcessed++
        Write-BackfillLog "[$graphProcessed/$($graphBatch.Count)] Setting Graph ShowInAddressList=$($item.DesiredShowInAddressList) for $($item.Email), CAPID $($item.CAPID), ObjectId $($item.User.id)"

        if ($DryRun) {
            continue
        }

        try {
            $updateUri = "https://graph.microsoft.com/beta/users/$($item.User.id)"
            $updateBody = @{
                showInAddressList = $item.DesiredShowInAddressList
            } | ConvertTo-Json
            Invoke-MgGraphRequest -Method PATCH -Uri $updateUri -Body $updateBody -ContentType "application/json"
            Write-BackfillLog "Updated Graph ShowInAddressList for $($item.Email)."
        } catch {
            if (-not $graphFailureCounts.ContainsKey($item.User.id)) {
                $graphFailureCounts[$item.User.id] = 0
            }
            $graphFailureCounts[$item.User.id]++
            Write-BackfillLog "Failed to update Graph ShowInAddressList for $($item.Email). Failure $($graphFailureCounts[$item.User.id]) of $MaxFailuresPerRecipient. Error: $($_.Exception.Message)"
        }
    }

    $toUpdate = @()
    if (-not $SkipExchange) {
        $hiddenGuestRecipients = @(Get-Recipient -Filter "RecipientTypeDetails -eq 'GuestMailUser' -and HiddenFromAddressListsEnabled -eq 'True'" -ResultSize Unlimited -ErrorAction Stop)
        Write-BackfillLog "Found $($hiddenGuestRecipients.Count) hidden guest mail recipient(s) in Exchange."

        foreach ($recipient in $hiddenGuestRecipients) {
            $matchedUser = Find-MatchedGuestUser -Recipient $recipient -GuestUsers $guestUsers
            if (-not $matchedUser) {
                continue
            }

            $capid = "$($matchedUser.employeeId)".Trim()
            if ($capid -match '(?i)P$') {
                continue
            }

            $identity = if ($recipient.Identity) { "$($recipient.Identity)" } else { "$($matchedUser.userPrincipalName)" }
            if ($failureCounts.ContainsKey($identity) -and $failureCounts[$identity] -ge $MaxFailuresPerRecipient) {
                continue
            }

            $toUpdate += [PSCustomObject]@{
                Recipient = $recipient
                User      = $matchedUser
                Identity  = $identity
                Email     = if ($matchedUser.mail) { $matchedUser.mail } elseif ($recipient.PrimarySmtpAddress) { $recipient.PrimarySmtpAddress } else { $matchedUser.userPrincipalName }
                CAPID     = $capid
            }
        }
    }

    if ($toUpdate.Count -eq 0 -and $graphUpdates.Count -eq 0) {
        Write-BackfillLog "No hidden non-parent guest account(s) remain to set visible."
        Write-BackfillLog "No guest Graph ShowInAddressList mismatch(es) remain."
        break
    }

    $batch = if ($BatchSize -gt 0) { @($toUpdate | Select-Object -First $BatchSize) } else { @($toUpdate) }
    if ($batch.Count -gt 0) {
        Write-BackfillLog "Updating $($batch.Count) of $($toUpdate.Count) hidden non-parent guest account(s) this pass."
    }

    $processed = 0
    foreach ($item in $batch) {
        $processed++
        Write-BackfillLog "[$processed/$($batch.Count)] Setting HiddenFromAddressListsEnabled=False for $($item.Email), CAPID $($item.CAPID), Identity $($item.Identity)"

        if ($DryRun) {
            continue
        }

        try {
            Set-MailUser -Identity $item.Identity -HiddenFromAddressListsEnabled $false -ErrorAction Stop
            Write-BackfillLog "Updated $($item.Email)."
        } catch {
            if (-not $failureCounts.ContainsKey($item.Identity)) {
                $failureCounts[$item.Identity] = 0
            }
            $failureCounts[$item.Identity]++
            Write-BackfillLog "Failed to update $($item.Email). Failure $($failureCounts[$item.Identity]) of $MaxFailuresPerRecipient. Error: $($_.Exception.Message)"
        }
    }

    if ($toUpdate.Count -le $batch.Count -and $graphUpdates.Count -le $graphBatch.Count) {
        Write-BackfillLog "Completed all currently detected hidden non-parent guest updates."
        Write-BackfillLog "Completed all currently detected Graph ShowInAddressList updates."
        continue
    }

    Write-BackfillLog "Sleeping $SleepSeconds seconds before next pass."
    Start-Sleep -Seconds $SleepSeconds
}

$failedRecipients = @($failureCounts.GetEnumerator() | Where-Object { $_.Value -ge $MaxFailuresPerRecipient })
$failedGraphUsers = @($graphFailureCounts.GetEnumerator() | Where-Object { $_.Value -ge $MaxFailuresPerRecipient })
if ($failedRecipients.Count -gt 0 -or $failedGraphUsers.Count -gt 0) {
    Write-BackfillLog "Finished with $($failedRecipients.Count) Exchange recipient(s) and $($failedGraphUsers.Count) Graph user(s) skipped after repeated failures."
    exit 1
}

Write-BackfillLog "Backfill complete."
