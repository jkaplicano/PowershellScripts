<#
.SYNOPSIS
    Sync AAD security group membership to on-prem AD security group.

.DESCRIPTION
    This script compares Azure AD group membership with an on-premises AD group
    and adds or removes members as needed to keep them in sync, using the Azure AD group as source.

.AUTHOR
    Amed Aplicano

.DATE
    11/02/2023
#>

$path = "Path..."
$Date = (Get-Date).ToString('MM-dd-yyyy_HH-mm')
$transcriptpath = "$path" + $Date + ".log" + $Date + ".log"

Start-Transcript -Path $transcriptpath -ErrorAction Continue -NoClobber

# Import modules
Import-Module azuread
Import-Module activedirectory

#Azure AD authentication
$password = "Password"
$username = 'UserName@contoso.com'
$Usercredential = New-Object System.Management.Automation.PSCredential($username,$password)

try {
    Write-Host "Attempting to connect to Azure AD."
    Connect-AzureAD -Credential $Usercredential
    Write-Host "Connected to Azure AD successfully."
}
catch {
    Write-Host "Failed to connect to Azure AD: $_" -ForegroundColor Red
    return  # Exit the script if we cannot connect
}

# Source CSV File
$sourcecsvFilePath = "Path...\FileName.csv"
try {
    Write-Host "Importing group pairs from CSV file."
    $groupPairs = Import-Csv -Path $sourcecsvFilePath
    Write-Host "Imported $($groupPairs.Count) group pairs from the CSV file."
}
catch {
    Write-Host "Error importing CSV file: $_" -ForegroundColor Red
    return  # Exit the script if CSV import fails
}

# Paths for added/removed logs
$AddedMembersLogsPath = 'Path...\MembersAdded.log'
$RemovedUsersLogsPath = 'Path...\MembersRemoved.log'

foreach ($groupPair in $groupPairs) {
    $SourceGroupName = $groupPair.Source
    $DestinationGroupName = $groupPair.Destination

    # Sync Start
    Write-Host "Starting sync for group '$SourceGroupName' to '$DestinationGroupName'."

    # Error handling for Source Group Members retrieval (Azure AD)
    try {
        Write-Host "Retrieving members of the Azure AD group '$SourceGroupName'."
        $SourceGroupUPNs = Get-AzureADGroupMember -ObjectId $SourceGroupName -All $true | 
            Where-Object { $_.UserPrincipalName -ne $null } | 
            Select-Object -ExpandProperty UserPrincipalName
        Write-Host "Retrieved $(($SourceGroupUPNs).Count) members from Azure AD group '$SourceGroupName'."
    }
    catch {
        Write-Host "Error retrieving source group UPNs for group '$SourceGroupName'. Error: $_" -ForegroundColor Red
        continue  
    }

    # Error handling for Destination Group Members retrieval (On-Prem AD)
    try {
        Write-Host "Retrieving members of the on-prem AD group '$DestinationGroupName'."
        $DestinationGroupDNs = Get-ADGroupMember -Identity $DestinationGroupName
        Write-Host "Retrieved $(($DestinationGroupDNs).Count) members from on-prem AD group '$DestinationGroupName'."
    }
    catch {
        Write-Host "Error retrieving destination group members for group '$DestinationGroupName'. Error: $_" -ForegroundColor Red
        continue  
    }

    # Initialize an empty array for UPNs from the destination group
    $DestinationGroupUPNs = @()

    # Error handling for retrieving UserPrincipalName from AD users
    try {
        foreach ($DN in $DestinationGroupDNs) {
            $ADUser = Get-ADUser -Identity $DN.DistinguishedName
            if ($ADUser.UserPrincipalName) {
                $DestinationGroupUPNs += $ADUser.UserPrincipalName
            }
        }
        Write-Host "Retrieved UPNs for $(($DestinationGroupUPNs).Count) members of '$DestinationGroupName'."
    }
    catch {
        Write-Host "Error retrieving UPNs for destination group members in '$DestinationGroupName'. Error: $_" -ForegroundColor Red
        continue  
    }

    # Create an empty HashSet for the source group UPNs
    Write-Host "Creating HashSet for source group UPNs."
    $SourceHashSet = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($UPN in $SourceGroupUPNs) {
        [void]$SourceHashSet.Add($UPN)
    }

    # Create an empty HashSet for the destination group UPNs
    Write-Host "Creating HashSet for destination group UPNs."
    $DestinationHashSet = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($UPN in $DestinationGroupUPNs) {
        [void]$DestinationHashSet.Add($UPN)
    }

    # Define Members to Add and Members to Remove here
    $MembersToAdd = $SourceGroupUPNs | Where-Object { -not $DestinationHashSet.Contains($_) }
    Write-Host "Found $(($MembersToAdd).Count) members to add to '$DestinationGroupName'."

    $MembersToRemove = $DestinationGroupUPNs | Where-Object { -not $SourceHashSet.Contains($_) }
    Write-Host "Found $(($MembersToRemove).Count) members to remove from '$DestinationGroupName'."

    # Adding members to the destination group
    if ($MembersToAdd.Count -gt 0) {
        try {
            Write-Host "Starting to add members to the destination group '$DestinationGroupName'."
            foreach ($MemberToAdd in $MembersToAdd) {
                $ADUser = Get-ADUser -Filter {UserPrincipalName -eq $MemberToAdd}
                if ($ADUser -eq $null) {
                    throw "User not found: $MemberToAdd"
                }
                $ADUserDisplayName = $ADUser.DisplayName
                $ADUserDN = $ADUser.DistinguishedName
                Write-Host "Adding $ADUserDisplayName to group $DestinationGroupName"
                Add-ADGroupMember -Identity $DestinationGroupName -Members $ADUserDN
                Add-Content -Path $AddedMembersLogsPath -Value "$(Get-Date) | $ADUserDisplayName | $MemberToAdd | $DestinationGroupName"
                Start-Sleep -Seconds 5
            }
        }
        catch {
            Write-Host "Error adding members to group '$DestinationGroupName'. Error: $_" -ForegroundColor Red
            continue  
        }
    } else {
        Write-Host "No new members to add to group $DestinationGroupName."
    }

    # Removing members from the destination group
    if ($MembersToRemove.Count -gt 0) {
        try {
            Write-Host "Starting to remove members from the destination group '$DestinationGroupName'."
            foreach ($MemberToRemove in $MembersToRemove) {
                $ADUser2 = Get-ADUser -Filter {UserPrincipalName -eq $MemberToRemove}
                if ($ADUser2 -eq $null) {
                    throw "User not found: $MemberToRemove"
                }
                $ADUser2DisplayName = $ADUser2.DisplayName
                $ADUser2DN = $ADUser2.DistinguishedName

                Write-Host "Removing $ADUser2DisplayName from group $DestinationGroupName" -ForegroundColor Black -BackgroundColor White
                Remove-ADGroupMember -Identity $DestinationGroupName -Members $ADUser2DN -Confirm:$false
                Add-Content -Path $RemovedUsersLogsPath -Value "$(Get-Date) | $ADUser2DisplayName | $MemberToRemove | $DestinationGroupName"
                Start-Sleep -Seconds 5
            }
        }
        catch {
            Write-Host "Error removing members from group '$DestinationGroupName'. Error: $_" -ForegroundColor Red
            continue  
        }
    } else {
        Write-Host "No members to remove from group $DestinationGroupName."
    }

    Write-Host "Next loop in 15 seconds" -ForegroundColor Yellow
    Start-Sleep -Seconds 15
}

# Disconnect from Azure AD and stop transcript
try {
    Write-Host "Disconnecting from Azure AD."
    Disconnect-AzureAD -Confirm:$false
}
catch {
    Write-Host "Error disconnecting from Azure AD: $_" -ForegroundColor Red
}

Stop-Transcript
