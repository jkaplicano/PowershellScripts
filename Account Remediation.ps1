<#
.SYNOPSIS
    Account Remediation

.DESCRIPTION
    This script remediates accounts by resetting passwords, revoking tokens, 
    and removing inbox rules for compromised users.

.AUTHOR
    Amed Aplicano

.DATE
    06/10/2023
#>

# Logs
$Date = (Get-Date).ToString('MM-dd-yyyy--HHmm')

$path = "Path..."
$transcriptpath = "$path" + $Date + ".log"
Start-Transcript -Path $transcriptpath -NoClobber -ErrorAction Continue

# Modules and Authentication

function Authenticate {
    param(
        [int]$maxRetries = 3
    )

    for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
        try {
            cls
            Set-ExecutionPolicy Unrestricted -f

            # Connect to Exchange Online Modern Auth
            Import-Module ExchangeOnlineManagement

            try {
                $password = "Password"
                $username = 'UserName@Contoso.com'
                $Usercredential = New-Object System.Management.Automation.PsCredential($username,$password)
                Connect-ExchangeOnline -Credential $Usercredential
            }
            catch {
                Write-Host "Error: Failed to connect to Exchange Online. $($_.Exception.Message)" -ForegroundColor Red
                return
            }

            # MS Graph Authentication
            try {
                $EncryptedData = Get-Content "Path...\AppSecret.encrypted"
            }
            catch {
                Write-Host "Error: Unable to read the encrypted data file. Please check the file path and permissions." -ForegroundColor Red
                return
            }

            try {
                $PasswordSecureString = ConvertTo-SecureString $EncryptedData
                $clientID = "ApplicationID"
                $Clientsecret = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($PasswordSecureString))
                $tenantID = "AzureADTenantID"

                $tokenBody = @{
                    Grant_Type    = "client_credentials"
                    Scope         = "https://graph.microsoft.com/.default"
                    Client_Id     = $clientID
                    Client_Secret = $clientSecret
                }

                $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token" -Method POST -Body $tokenBody

                # Check if token is null or missing
                if (-not $tokenResponse.access_token) {
                    Write-Host "Error: Failed to retrieve access token from Graph API." -ForegroundColor Red
                    return
                }

                $headersGraph = @{
                    "Authorization" = "Bearer $($tokenResponse.access_token)"
                    "Content-type"  = "application/json"
                }
            }
            catch {
                Write-Host "Error: Failed to authenticate with the Microsoft Graph API. Error details: $($_.Exception.Message)" -ForegroundColor Red
                return
            }

            Write-Host "Exchange Online connection succeeded" -ForegroundColor DarkYellow
            Start-Sleep -Seconds 5

            # Additional connections
            Import-Module ActiveDirectory
            Import-Module MSOnline

            try {
                Connect-MsolService -Credential $UserCredential
                Write-Host "MSOL connection succeeded" -ForegroundColor DarkYellow
            }
            catch {
                Write-Host "Error: Failed to connect to MSOL service. Please check credentials or network connectivity. Error: $($_.Exception.Message)" -ForegroundColor Red
                return
            }

            try {
                Connect-AzureAD -Credential $UserCredential
                Write-Host "AzureAD connection succeeded" -ForegroundColor DarkYellow
            }
            catch {
                Write-Host "Error: Failed to connect to AzureAD service. Please check credentials or network connectivity. Error: $($_.Exception.Message)" -ForegroundColor Red
                return
            }

            break  # Exit loop if successful
        }
        catch {
            Write-Host "Authentication error on attempt $($attempt): $_" -ForegroundColor Red

            if ($attempt -lt $maxRetries) {
                Start-Sleep -Seconds 5  # Wait before retrying
            }
            else {
                Write-Host "Maximum number of retries reached. Exiting." -ForegroundColor Red
                return
            }
        }
    }
}

# Call the Authenticate function with a maximum of 3 retries
Authenticate 

# General Variables

$affectedUPN = Read-Host "Enter Affected User UPN"
$sender = "UserName@Contoso.com"
$target = "UserName@Contoso.com"
$parts = $affectedUPN.Split('@')
$username = $parts[0]
$UserMailbox = "UserMailbox"
$SharedMailbox = "SharedMailbox"

# Passphrase password

try {
    $wordsfile = "Path...\words.txt"
    $wordsList = Get-Content $wordsfile
    $passphraseWords = Get-Random -InputObject $wordsList -Count 3
    $passphrase = $passphraseWords -join "-"
    $passphrase1 = ConvertTo-SecureString $passphrase -AsPlainText -Force
    Write-Host "Passphrase Password: $passphrase"
}
catch {
    Write-Host "Error: Unable to generate or read passphrase. Please check file paths and permissions." -ForegroundColor Red
    return
}

# Actions

try {
    $mailboxobject = get-mailbox -identity $affectedUPN
    $mailboxtype = $mailboxobject.RecipientTypeDetails
    $Azureadobject = Get-AzureADUser -ObjectId $affectedUPN
    $azureadobjectid = $Azureadobject.ObjectId
    $azureadupn = $Azureadobject.UserPrincipalName
}
catch {
    Write-Host "Error: Failed to retrieve mailbox or Azure AD object. Please verify UPN. $($_.Exception.Message)" -ForegroundColor Red
    return
}

# Delegates

try {
    $FullAccessDelegates = Get-MailboxPermission -Identity $affectedUPN -ResultSize unlimited | Where-Object {($_.IsInherited -ne "True") -and ($_.User -notlike "*SELF*")}
    $SendAsDelegates = Get-Mailbox $affectedUPN | Get-RecipientPermission -ResultSize unlimited | Where-Object {($_.trustee -ne "NT AUTHORITY\SELF")}
}
catch {
    Write-Host "Error: Failed to retrieve mailbox delegates. $($_.Exception.Message)" -ForegroundColor Red
    return
}

$difference = $SendAsDelegates | Select @{N='User';E={$_.Trustee}},@{N='AccessRights';E={$_.AccessRights}}
$reference = $FullAccessDelegates | Select @{N='User';E={$_.user}},@{N='AccessRights';E={$_.AccessRights}}

$alldelegates = $reference + $difference

# User/SharedMailbox Remediation
if ($mailboxtype -eq $UserMailbox -and $affectedUPN -eq $azureadupn) {
    Write-Host "This is a user Mailbox" -BackgroundColor White -ForegroundColor Black
    Start-Sleep -Seconds 5

    try {
        Write-Host "Revoking AAD Sessions Token" -BackgroundColor White -ForegroundColor Black
        Revoke-AzureADUserAllRefreshToken -ObjectId $azureadobjectid
        Start-Sleep -Seconds 5

        Write-Host "Changing Password" -BackgroundColor White -ForegroundColor Black
        Set-AzureADUserPassword -ObjectId $azureadobjectid -Password $passphrase1
        Set-ADAccountPassword -Identity $username -NewPassword $passphrase1
        Set-MsolUser -UserPrincipalName $affectedUPN -StrongPasswordRequired $True
        Start-Sleep -Seconds 5
    }
    catch {
        Write-Host "Error: Failed to reset user password or revoke session tokens. $($_.Exception.Message)" -ForegroundColor Red
        return
    }

    Write-Host "Checking for Inbox Rules and Mail Forwarding" -BackgroundColor White -ForegroundColor Black
    try {
        Get-InboxRule -Mailbox $affectedUPN | Where-Object { $_.Enabled -and ($_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo -or $_.SendTextMessageNotificationTo) } | Format-Table
        Get-InboxRule -Mailbox $affectedUPN | Where-Object { $_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo -or $_.SendTextMessageNotificationTo } | Remove-InboxRule -Confirm:$false
        Get-InboxRule -Mailbox $affectedUPN | Where-Object { $_.Enabled -and ($_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo -or $_.SendTextMessageNotificationTo) } | Disable-InboxRule -Confirm:$false
        Get-Mailbox -Identity $affectedUPN | Select Name, DeliverToMailboxAndForward, ForwardingSmtpAddress
        Set-Mailbox -Identity $affectedUPN -DeliverToMailboxAndForward $false -ForwardingSmtpAddress $null
        Start-Sleep -Seconds 5
    }
    catch {
        Write-Host "Error: Failed to remove inbox rules or forwarding addresses. $($_.Exception.Message)" -ForegroundColor Red
    }

    Get-Mailbox -Identity $affectedUPN | Select Name, DeliverToMailboxAndForward, ForwardingSmtpAddress
} else {
    if ($mailboxtype -eq $SharedMailbox -and $affectedUPN -eq $azureadupn) {
        Write-Host "This is a SharedMailbox" -BackgroundColor White -ForegroundColor Black
        Start-Sleep -Seconds 5

        try {
            Write-Host "Revoking AAD Sessions Token" -BackgroundColor Red -ForegroundColor Yellow
            Revoke-AzureADUserAllRefreshToken -ObjectId $azureadobjectid
            Start-Sleep -Seconds 5

            Write-Host "Changing Password" -BackgroundColor Red -ForegroundColor Yellow
            Set-AzureADUserPassword -ObjectId $azureadobjectid -Password $passphrase1
            Set-ADAccountPassword -Identity $username -NewPassword $passphrase1
            Set-MsolUser -UserPrincipalName $affectedUPN -StrongPasswordRequired $True
            Start-Sleep -Seconds 5
        }
        catch {
            Write-Host "Error: Failed to reset shared mailbox password or revoke session tokens. $($_.Exception.Message)" -ForegroundColor Red
            return
        }

        Write-Host "Checking for Inbox Rules and Mail Forwarding" -BackgroundColor Red -ForegroundColor Yellow
        try {
            Get-InboxRule -Mailbox $affectedUPN | Where-Object { $_.Enabled -and ($_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo -or $_.SendTextMessageNotificationTo) } | Format-Table
            Get-InboxRule -Mailbox $affectedUPN | Where-Object { $_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo -or $_.SendTextMessageNotificationTo } | Remove-InboxRule -Confirm:$false
            Get-InboxRule -Mailbox $affectedUPN | Where-Object { $_.Enabled -and ($_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo -or $_.SendTextMessageNotificationTo) } | Disable-InboxRule -Confirm:$false
            Get-Mailbox -Identity $affectedUPN | Select Name, DeliverToMailboxAndForward, ForwardingSmtpAddress
            Set-Mailbox -Identity $affectedUPN -DeliverToMailboxAndForward $false -ForwardingSmtpAddress $null
            Start-Sleep -Seconds 5
        }
        catch {
            Write-Host "Error: Failed to remove inbox rules or forwarding addresses from shared mailbox. $($_.Exception.Message)" -ForegroundColor Red
        }

        Get-Mailbox -Identity $affectedUPN | Select Name, DeliverToMailboxAndForward, ForwardingSmtpAddress

        foreach ($delegate in $alldelegates) {
            try {
                Write-Host "Removing delegation and revoking active sessions for $($delegate.User)"
                Remove-MailboxPermission -Identity $affectedUPN -User $delegate.User -AccessRights $delegate.AccessRights -InheritanceType All -Confirm:$false
                $delegateobject = Get-AzureADUser -ObjectId $delegate.User
                Revoke-AzureADUserAllRefreshToken -ObjectId $delegateobject.ObjectId
                Start-Sleep -Seconds 10
            }
            catch {
                Write-Host "Error: Failed to remove delegation for $($delegate.User). $($_.Exception.Message)" -ForegroundColor Red
            }
        }

        foreach ($delegate in $alldelegates) {
            try {
                Write-Host "Re-assigning delegation for $($delegate.User)"
                Add-MailboxPermission -Identity $affectedUPN -User $delegate.User -AccessRights $delegate.AccessRights -InheritanceType All
                Start-Sleep -Seconds 10
            }
            catch {
                Write-Host "Error: Failed to re-assign delegation for $($delegate.User). $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    } else {
        Write-Host "This is not a User or Shared Mailbox, please try again."
    }
}

# Running Mailbox Auditing Process (7 Days)
try {
    $auditLogPath = "$path" + $userName + "_" + "AuditLog" + $date + ".csv"
    $startDate = (Get-Date).AddDays(-7).ToString('MM/dd/yyyy') 
    $endDate = (Get-Date).ToString('MM/dd/yyyy')
    $results = Search-MailboxAuditLog -Identity $affectedUPN -ShowDetails -StartDate $startDate -EndDate $endDate | Select-Object Operation, OperationResult, LogonType, LogonUserDisplayName, ClientProcessName, ItemSubject, LastAccessed
    $results | Export-Csv -Path $auditLogPath -NoTypeInformation
}
catch {
    Write-Host "Error: Failed to retrieve or export mailbox audit logs. $($_.Exception.Message)" -ForegroundColor Red
}

# Create Samanage Ticket
try {
    Add-Type -AssemblyName System.Web

    $ADobject = Get-ADUser $username -Properties physicalDeliveryOfficeName, Displayname | Select-Object Displayname, physicalDeliveryOfficeName
    $Displayname = $ADobject.Displayname
    $Usersite = $ADobject.physicalDeliveryOfficeName
    $UserITFEgroup = "!!" + $Usersite + "ITFE"

    $JsonWebToken = "U1dTRC1pZG1BUElAZHByLmNvbQ==:eyJhbGciOiJIUzUxMiJ9.eyJ1c2VyX2lkIjo1Njk2Mzc5LCJnZW5lcmF0ZWRfYXQiOiIyMDIwLTA0LTI3IDE5OjE2OjAzIn0.Ndvz5eaeD1AFhl25kMh-8-buw0BuMUm1EhS9OWBYyNrCrM9KEnXaiAIH1z3Rx_LZHrlLrHC0IoBsZl6wc7qFMg"
    $URI = "https://api.samanage.com/incidents.xml"

    $headers = @{
        "X-Samanage-Authorization" = "Bearer $JsonWebToken"
        "Accept" = "application/vnd.samanage.v2.1+xml"
        "Content-Type" = "application/xml"
    }

    $incidentName = "Compromised Account - $Displayname"

    $HTMLFile = @"
<incident>
    <name>$incidentName</name>
    <priority>High</priority>
    <requester><email>SWSD-idmAPI@dpr.com</email></requester>
    <custom_fields_values>
        <custom_fields_value>
            <name>Threat reported</name>
            <value>Compromised Account</value>
        </custom_fields_value>
    </custom_fields_values>
    <category><name>Cybersecurity</name></category>
    <subcategory><name>Compromised Account</name></subcategory>
    <description>
        Affected User: $username
        User email: $affectedUPN
        New Temporary Password: $passphrase
        Site: $Usersite
        ITFE Routing Automation: $UserITFEgroup
    </description>
    <due_at>Dec 11, 2021</due_at>
    <assignee><email>samanage_api@dpr.com</email></assignee>
</incident>
"@

    Write-Host "Search Samanage Incident"

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    Invoke-RestMethod -Body $HTMLFile -Method POST -Headers $Headers -Uri $URI

    $URIJSON = "https://api.samanage.com/incidents.json"

    $HeadersJSON = @{
        "X-Samanage-Authorization" = "Bearer $JsonWebToken"
        "Accept" = "application/vnd.samanage.v2.1+json"
        "Content-Type" = "application/json"
    }

    $waitDuration = 10
    Start-Sleep -Seconds $waitDuration

    $holdResponse = Invoke-RestMethod -Method Get -Headers $HeadersJSON -Uri $URIJSON

    $incidentName = "Compromised Account - $Displayname"
    $index = $holdResponse.name.IndexOf($incidentName)

    if ($index -lt 0) {
        Write-Host "Error: Incident not found." -ForegroundColor Red
        return
    }

    $IncidentNumber = $holdResponse.number[$index]
    $IncidentURL_ID = $holdResponse.id[$index]
    $URL = "https://Contoso.samanage.com/incidents/$IncidentURL_ID"

    Write-Host "Samanage INCIDENT # $IncidentNumber" -ForegroundColor Yellow
    Write-Host "Samanage URL Link $URL"
}
catch {
    Write-Host "Error: Failed to create or retrieve Samanage incident. $($_.Exception.Message)" -ForegroundColor Red
}

# Incremental Log

try {
    New-Object -TypeName PSCustomObject -Property @{
        "Date"               = $Date
        "Samanage Ticket"     = $IncidentNumber
        "Username"           = $userName
        "PrimarySMTPAddress"  = $affectedUPN
        "Audit Log"          = $auditLogPath
        "Content Search"     = ""
        "MS Investigation"   = ""
    } | Export-Csv "$path\AccountRemediationProcessIncremental_Log.csv" -NoTypeInformation -Append
}
catch {
    Write-Host "Error: Failed to update incremental log. $($_.Exception.Message)" -ForegroundColor Red
}

# Disconnect services
if (Get-Module -ListAvailable -Name AzureAD) {
    Disconnect-AzureAD -Confirm:$false
}

if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
    Disconnect-ExchangeOnline -Confirm:$false
}
