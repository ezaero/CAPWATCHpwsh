# Input bindings are passed in via param block.
param($Timer)

# Set working directory to folder with all CAPWATCH CSV Text Files
$CAPWATCHDATADIR = "$($env:HOME)\data\CAPWatch"
Push-Location $CAPWATCHDATADIR

# Include shared Functions
. "$PSScriptRoot\..\shared\shared.ps1"

# Connect to Microsoft Graph
$MSGraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token

Connect-MgGraph -AccessToken $MSGraphAccessToken -NoWelcome
Connect-ExchangeOnline -ManagedIdentity -Organization $env:EXCHANGE_ORGANIZATION

# Helper function to get notification emails for a unit's commanders and recruiting officer
function Get-UnitNotificationEmails {
    param (
        [string]$unit,
        [array]$allUsers
    )
    $emails = @()
    foreach ($user in $allUsers) {
        if ($user.companyName -match $unit -and $user.department -match '(PA|EX)') {
            if ($user.mail) { $emails += $user.mail }
        }
    }
    $emails = $emails | Select-Object -Unique
    return $emails
}

# Function to send expired members notification email
function Send-ExpiredMembersNotification {
    param (
        [array]$deletedMembers,
        [array]$allUsers
    )
    
    if ($deletedMembers.Count -eq 0) {
        Write-Log "No expired members to report."
        return
    }
    
    # Group deleted members by unit
    $deletedByUnit = $deletedMembers | Group-Object -Property Unit
    
    foreach ($unitGroup in $deletedByUnit) {
        $unit = $unitGroup.Name
        $unitMembers = $unitGroup.Group
        
        # Get notification emails for this unit
        $unitEmails = Get-UnitNotificationEmails -unit $unit -allUsers $allUsers
        if ($unitEmails.Count -eq 0) {
            Write-Log "No notification emails found for unit $unit. Skipping notification."
            continue
        }
        
        # Always include mike.schulte@cowg.cap.gov in the recipients
        $toRecipients = @('mike.schulte@cowg.cap.gov')
        foreach ($email in $unitEmails) {
            if ($email -and $email -ne 'mike.schulte@cowg.cap.gov') {
                $toRecipients += $email
            }
        }
        
        # Build the member list HTML table
        $memberTableRows = ""
        foreach ($member in $unitMembers) {
            $memberTableRows += @"
      <tr>
        <td style='padding: 8px; border: 1px solid #ddd;'>$($member.NameFirst) $($member.NameLast)</td>
        <td style='padding: 8px; border: 1px solid #ddd;'>$($member.Grade)</td>
        <td style='padding: 8px; border: 1px solid #ddd;'>$($member.CAPID)</td>
        <td style='padding: 8px; border: 1px solid #ddd;'>$($member.Email)</td>
      </tr>
"@
        }
        
        try {
            $userPrincipalName = "cowg_it_helpdesk@cowg.cap.gov" # Use a service account or shared mailbox with Mail.Send permission
            $mailBody = @{
                message = @{
                    subject = "Expired Members Removed from CO-$unit"
                    body = @{
                        contentType = "HTML"
                        content = @"
<html>
  <body style='font-family: Arial, sans-serif; color: #222;'>
    <div style='text-align: center; margin-bottom: 20px;'>
      <img src='https://cowg.cap.gov/media/websites/COWG_T_7665FADF8B38C.PNG' alt='COWG Logo' style='max-width: 200px;'/>
    </div>
    <h2 style='color: #003366;'>Expired Members Removed from CO-$unit</h2>
    <p>The following members have been removed from CO-$unit because their membership has expired in CAPWATCH:</p>
    <table style='margin: 20px auto; border-collapse: collapse; width: 90%;'>
      <thead>
        <tr style='background-color: #f2f2f2;'>
          <th style='padding: 8px; border: 1px solid #ddd; text-align: left;'>Name</th>
          <th style='padding: 8px; border: 1px solid #ddd; text-align: left;'>Grade</th>
          <th style='padding: 8px; border: 1px solid #ddd; text-align: left;'>CAPID</th>
          <th style='padding: 8px; border: 1px solid #ddd; text-align: left;'>Email</th>
        </tr>
      </thead>
      <tbody>
$memberTableRows
      </tbody>
    </table>
    <p><strong>Total removed:</strong> $($unitMembers.Count) member(s)</p>
    <p style='font-size: 0.9em; color: #888; margin-top: 30px;'>This is an automated notification from the COWG IT Team. These accounts have been permanently deleted from Azure AD and Exchange Online.</p>
  </body>
</html>
"@
                    }
                    toRecipients = $toRecipients
                }
                saveToSentItems = $false
            } | ConvertTo-Json -Depth 4
            
            $uri = "https://graph.microsoft.com/v1.0/users/$userPrincipalName/sendMail"
            Invoke-MgGraphRequest -Method POST -Uri $uri -Body $mailBody -ContentType "application/json"
            Write-Log "Expired members notification email sent for unit CO-$unit to: $($toRecipients -join ', ')"
        } catch {
            Write-Log "Failed to send expired members notification email for unit CO-${unit}: $_"
        }
    }
}

# Main execution block
try {
    Write-Log "Starting maintenance operations..."
    
    # Clean up old log files
    Remove-OldLogFiles -DirectoryPath "$env:HOME\logs"
    
    # Run monthly account deletion maintenance
    Write-Log "Running monthly account deletion maintenance..."
    Remove-ExpiredMemberAccounts
    
    Write-Log "Maintenance operations completed successfully."
} catch {
    Write-Log "Error during maintenance operations: $_"
    throw
}
