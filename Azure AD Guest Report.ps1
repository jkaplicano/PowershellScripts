<#
.SYNOPSIS
    Guest Accounts Report

.DESCRIPTION
    This script generates a CSV report of guest accounts in Azure AD tenant and searches for last activity in MS Graph Audit Logs
.AUTHOR
    Amed Aplicano

.DATE
    
#>
$MaximumFunctionCount = 16384
import-module Microsoft.Graph.Reports
import-module azuread

# MS Graph Authentication #
$appid = 'ApplicationID'
$tenantid = 'TenantID'
$EncryptedData = Get-Content “Path...\Secret.encrypted”

try {
    $PasswordSecureString = ConvertTo-SecureString $EncryptedData
    $secretID = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($PasswordSecureString))
    $secret = "$secretID"
} catch {
    Write-Error "Error occurred while processing the secret: $_"
    exit 1
}

$body = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    Client_Id     = $appid
    Client_Secret = $secret
}

# Error Handling for API Request
try {
    $connection = Invoke-RestMethod `
        -Uri https://login.microsoftonline.com/$tenantid/oauth2/v2.0/token `
        -Method POST `
        -Body $body
} catch {
    Write-Error "Failed to authenticate with Microsoft Graph: $_"
    exit 1
}

$token = $connection.access_token

if (-not $token) {
    Write-Error "Failed to retrieve access token."
    exit 1
}

$secureToken = ConvertTo-SecureString -AsPlainText $token -Force

try {
    Connect-MgGraph -AccessToken $secureToken
} catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    exit 1
}

# User credentials securely input (you can update to use prompts)
$password = Read-Host -AsSecureString "Enter your password"
$username = 'UserName@Contoso.com'
$userCredential = New-Object System.Management.Automation.PSCredential($username, $password)

# Azure AD Authentication
try {
    Connect-AzureAD -Credential $userCredential
} catch {
    Write-Error "Azure AD Authentication failed: $_"
    exit 1
}

# Fetch guest users with refactored conditions
try {
    $allguest = Get-AzureADUser -All:$true | Where-Object {
        $_.UserType -eq 'Guest' -and ($_.UserPrincipalName -like '*')
    }
} catch {
    Write-Error "Error retrieving guest users: $_"
    exit 1
}

# Select guest properties
$filteredGuest = $allguest | Select-Object -Property AccountEnabled, UserType, DisplayName, UserPrincipalName, Mail, UserState

# Initialize array for report collection
$guestReports = @()

# Process each guest
foreach ($guest in $filteredGuest) {
    $UPN = $guest.Mail
    Write-Host 'Now Processing' $UPN

    try {
        $Log = Get-MgAuditLogSignIn -Filter "startsWith(UserPrincipalName, '$UPN')" -Top 1 | Select-Object ClientAppUsed, CreatedDateTime
    } catch {
        Write-Error "Error retrieving audit logs for $UPN: $_"
        continue
    }

    $guestReports += [pscustomobject]@{
        "UserType"        = $guest.UserType
        "Enabled"         = $guest.AccountEnabled
        "Display Name"    = $guest.DisplayName
        "GE Email Address"= $guest.Mail
        "Guest UPN"       = $guest.UserPrincipalName
        "Invitation"      = $guest.UserState
        "Last Activity"   = $Log.CreatedDateTime
        "ClientApp"       = $Log.ClientAppUsed
    }
}

# Export collected reports to CSV
$outputPath = "C:\Temp\GEGuestReport_Nov14.csv"
try {
    $guestReports | Export-Csv -Path $outputPath -NoTypeInformation
    Write-Host "Report successfully exported to $outputPath"
} catch {
    Write-Error "Failed to export report to CSV: $_"
}
