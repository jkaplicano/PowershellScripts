<#
.SYNOPSIS
    Remove former Employee Accounts from Licensing Security Group

.DESCRIPTION
    To identify terminated user objects with a license then remove them from the Licensing Groups.
.AUTHENTICATION

    AAD
    Exchange
    
.AUTHOR
    Amed Aplicano

.DATE
    March 15, 2024
#>

# Service Account Credentials

$password = "Password"
$username = 'UserName@Contoso.com'
$Usercredential = New-Object System.Management.Automation.PsCredential($username, $password)

# Group Variables
$GroupName = "GroupName"
$SourceAADGroup = 'ObjectID'

$RemovedUsersLogsPath = 'Path...\RemovedUsers.log'
$termMembersRemoved = 0
$count = 0

# Error Handling in Try-Catch Blocks
try {
    # AzureAD     Import-Module AzureAD
    Connect-AzureAD -Credential $UserCredential
} catch {
    Write-Host "Failed to connect to AzureAD: $_" -ForegroundColor Red
    exit
}

try {
    # Exchange Online 
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -Credential $Usercredential
} catch {
    Write-Host "Failed to connect to ExchangeOnline: $_" -ForegroundColor Red
    exit
}

# Get Group Members List
try {
    $SourceAADGroupMembersObjectID = (Get-AzureADGroupMember -All:$true -ObjectId $SourceAADGroup).ObjectId
    ($SourceAADGroupMembersObjectID).count
} catch {
    Write-Host "Failed to retrieve group members: $_" -ForegroundColor Red
    exit
}

# Loop to check user status and remove license
foreach ($member in $SourceAADGroupMembersObjectID) {
    try {
        $object = Get-EXOMailbox -Identity $member -Properties CustomAttribute13
        $EmployeeStatus = $object.CustomAttribute13
        $objectID = $object.ExternalDirectoryObjectId
        $AD = $object.Alias
        $UserLog = $object.UserPrincipalName

        if ($EmployeeStatus -eq 'A') { continue }

        # Check Active Directory Employee Status
        $EmployeeStatusAD = (Get-ADUser $AD -Properties extensionAttribute13).extensionAttribute13

        if ($EmployeeStatusAD -eq "T") {
            # Remove user from group and log the removal
            Remove-AzureADGroupMember -ObjectId $SourceAADGroup -MemberId $objectID
            $logEntry = [pscustomobject]@{
                DateTime = Get-Date
                User     = $UserLog
                Status   = 'Removed'
            }
            $logEntry | Export-Csv -Path $RemovedUsersLogsPath -Append -NoTypeInformation
            $termMembersRemoved++
        }
        
    } catch {
        Write-Host "Failed to process member ${member}: $_" -ForegroundColor Yellow
        continue
    }

    $count++
    Write-Host "Processed $count members"
}

Write-Host "Members removed: $termMembersRemoved"

# Disconnect in Finally Block
try {
    # Any final code that might need execution
} finally {
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-AzureAD -Confirm:$false
}
