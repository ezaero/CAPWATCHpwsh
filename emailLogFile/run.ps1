# Input bindings are passed in via param block.
param($Timer)

# Include shared Functions
 . "$PSScriptRoot\..\shared\shared.ps1"

# Connect to Microsoft Graph
$MSGraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token
Connect-MgGraph -AccessToken $MSGraphAccessToken -NoWelcome

# Define email parameters
$toAddress = "mike.schulte@cowg.cap.gov"
$fromAddress = "cowg_it_helpdesk@cowg.cap.gov"  # Replace with the app's email or a valid sender
$subject = "Daily Log File"
$body = "Here's the CAPWATCH Powershell log file for today: $(Get-Date -Format 'yyyy-MM-dd')."
$logFilePath = "$env:HOME\logs\script_log_$(Get-Date -Format 'yyyy-MM-dd').txt"

Write-Log "Sending email with log file for today's date: $(Get-Date -Format 'yyyy-MM-dd')"

# Check if the log file exists
if (-Not (Test-Path $logFilePath)) {
    Write-Host "Log file not found at $logFilePath"
    exit 1
}

# Read the log file content
$logFileContent = [System.IO.File]::ReadAllBytes($logFilePath)
$logFileBase64 = [System.Convert]::ToBase64String($logFileContent)

# Create the email payload
$emailPayload = @{
    message = @{
        subject = $subject
        body = @{
            contentType = "Text"
            content = $body
        }
        toRecipients = @(
            @{
                emailAddress = @{
                    address = $toAddress
                }
            }
        )
        attachments = @(
            @{
                "@odata.type" = "#microsoft.graph.fileAttachment"
                name = "script_log.txt"
                contentBytes = $logFileBase64
                contentType = "text/plain"
            }
        )
    }
    saveToSentItems = $true
} | ConvertTo-Json -Depth 10 -Compress

# Send the email using Microsoft Graph API
$userPrincipalName = $fromAddress  # The sender's email address
$uri = "https://graph.microsoft.com/v1.0/users/$userPrincipalName/sendMail"
Invoke-MgGraphRequest -Method POST -Uri $uri -Body $emailPayload -ContentType "application/json"

Write-Host "Email sent successfully to $toAddress"