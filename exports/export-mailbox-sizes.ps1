$ErrorActionPreference = "Stop"

Import-Module ExchangeOnlineManagement

$Organization = "cowg.cap.gov"
$OutputDir = Join-Path (Get-Location) "exports"
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$Stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$CsvPath = Join-Path $OutputDir "mailbox-sizes-$Stamp.csv"

Write-Host "Getting Exchange Online token..."
$Token = az account get-access-token --resource https://outlook.office365.com --query accessToken --output tsv

Write-Host "Connecting to Exchange Online..."
Connect-ExchangeOnline -AccessToken $Token -Organization $Organization -ShowBanner:$false

try {
    Write-Host "Enumerating mailboxes..."
    $Mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -Properties DisplayName,UserPrincipalName,PrimarySmtpAddress,RecipientTypeDetails,ArchiveStatus,WhenMailboxCreated,ExternalDirectoryObjectId)
    Write-Host "Found $($Mailboxes.Count) mailboxes. Writing $CsvPath"

    $Rows = New-Object System.Collections.Generic.List[object]
    $Index = 0

    foreach ($Mailbox in $Mailboxes) {
        $Index++
        if ($Index -eq 1 -or $Index % 25 -eq 0 -or $Index -eq $Mailboxes.Count) {
            Write-Host "Processing mailbox $Index of $($Mailboxes.Count): $($Mailbox.PrimarySmtpAddress)"
        }

        try {
            $Stats = Get-EXOMailboxStatistics -Identity $Mailbox.ExternalDirectoryObjectId
            $SizeString = [string]$Stats.TotalItemSize
            $Bytes = $null

            if ($SizeString -match "\(([-0-9,]+) bytes\)") {
                $Bytes = [double](($Matches[1]) -replace ",", "")
            }

            $Rows.Add([pscustomobject]@{
                DisplayName          = $Mailbox.DisplayName
                UserPrincipalName    = $Mailbox.UserPrincipalName
                PrimarySmtpAddress   = [string]$Mailbox.PrimarySmtpAddress
                RecipientTypeDetails = [string]$Mailbox.RecipientTypeDetails
                ArchiveStatus        = [string]$Mailbox.ArchiveStatus
                WhenMailboxCreated   = $Mailbox.WhenMailboxCreated
                ItemCount            = $Stats.ItemCount
                DeletedItemCount     = $Stats.DeletedItemCount
                TotalItemSize        = $SizeString
                SizeBytes            = $Bytes
                SizeGB               = if ($null -ne $Bytes) { [math]::Round($Bytes / 1GB, 3) } else { $null }
                LastLogonTime        = $Stats.LastLogonTime
                Error                = $null
            })
        }
        catch {
            $Rows.Add([pscustomobject]@{
                DisplayName          = $Mailbox.DisplayName
                UserPrincipalName    = $Mailbox.UserPrincipalName
                PrimarySmtpAddress   = [string]$Mailbox.PrimarySmtpAddress
                RecipientTypeDetails = [string]$Mailbox.RecipientTypeDetails
                ArchiveStatus        = [string]$Mailbox.ArchiveStatus
                WhenMailboxCreated   = $Mailbox.WhenMailboxCreated
                ItemCount            = $null
                DeletedItemCount     = $null
                TotalItemSize        = $null
                SizeBytes            = $null
                SizeGB               = $null
                LastLogonTime        = $null
                Error                = $_.Exception.Message
            })
        }
    }

    $SortedRows = $Rows | Sort-Object @{ Expression = "SizeBytes"; Descending = $true }
    $SortedRows | Export-Csv -NoTypeInformation -Path $CsvPath

    $Summary = [pscustomobject]@{
        MailboxCount = $Rows.Count
        CsvPath      = $CsvPath
        TotalSizeGB  = [math]::Round((($Rows | Measure-Object SizeBytes -Sum).Sum / 1GB), 3)
        ErrorCount   = @($Rows | Where-Object Error).Count
    }

    $Summary | ConvertTo-Json -Depth 4
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
