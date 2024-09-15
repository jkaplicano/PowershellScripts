<#
.SYNOPSIS
    New Conference Room Creation

.DESCRIPTION
    This script automates the conference room creation process, calendar processing settings, Delegations and location properties.

.AUTHOR
    Amed Aplicano

.DATE
    March 20, 2024
#>
# Parameters
Param (
    [parameter(Mandatory)]
    [ValidateSet('S', 'M')] 
    [string]$NumberofRooms
)

# Set Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
Unblock-File -Path "Path...\ConferenceRoomsFunctions.ps1" -Confirm:$false
Unblock-File -Path "Path...\AuthenticationFunctions.ps1" -Confirm:$false

# Load Functions
. "Path...\ConferenceRoomsFunctions.ps1"
. "Path...\AuthenticationFunctions.ps1"

# Log Path
$LogPath = "Path...\RoomsCreated.csv"

# Import Locations Table
try {
    $LocTable = ImportLocationsTable
    if ($LocTable -eq $null) { throw "Failed to import Locations Table" }
} catch {
    Write-Host "Error importing Locations Table: $_" -ForegroundColor Red
    return
}

# Single or Multiple Rooms
if ($NumberofRooms -eq "S") {
    $NewRoom = SingleRoomInfo
    if ($NewRoom -eq $null) { return }  # Exit if failed
    $NewRoom | FT
} elseif ($NumberofRooms -eq "M") {
    try {
        $Rooms = ImportRoomsInfoCSV
        Write-Host "Number of Rooms: $($Rooms.Count)" -ForegroundColor Yellow -BackgroundColor Black
        $Rooms | FT
    } catch {
        Write-Host "Error importing rooms CSV: $_" -ForegroundColor Red
        return
    }
} else {
    Write-Host "Invalid Parameter, Please Start Again" -ForegroundColor Red
    return
}

# Authentication
try {
    Write-Host "Authenticating to Exchange Online" -ForegroundColor Yellow -BackgroundColor Black
    ExchangeOnlineAuthenticationSchedulerQA
} catch {
    Write-Host "Error authenticating to Exchange Online: $_" -ForegroundColor Red
    return
}

try {
    Write-Host "Authenticating to O365" -ForegroundColor Yellow -BackgroundColor Black
    MSOLAuthenticationSchedulerQA
} catch {
    Write-Host "Error authenticating to O365: $_" -ForegroundColor Red
    return
}

# Room Account Password
$secret = RoomPassword
if ($secret -eq $null) { return }
$PWD = ConvertTo-SecureString $secret

# Room Creation Process
if ($NumberofRooms -eq "S") {
    ProcessRoom -Room $NewRoom -PWD $PWD -LocTable $LocTable -LogPath $LogPath
} elseif ($NumberofRooms -eq "M") {
    foreach ($Room in $Rooms) {
        ProcessRoom -Room $Room -PWD $PWD -LocTable $LocTable -LogPath $LogPath
    }
}

# Disconnect and Clean Up
Disconnect-ExchangeOnline -Confirm:$false
[Microsoft.Online.Administration.Automation.ConnectMsolService]::ClearUserSessionState()
