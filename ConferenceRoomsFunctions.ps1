Function ImportLocationsTable {
    Write-Host "Importing Locations Table" -ForegroundColor Yellow -BackgroundColor Black
    Start-Sleep 5

    try {
        $locations = Import-Csv -Path "Path...\RoomsLocations.csv"
        return $locations
    } catch {
        Write-Host "Error importing Locations Table: $_" -ForegroundColor Red
        return $null
    }
}

Function RoomPassword {
    Write-Host "Importing Room Password" -ForegroundColor Yellow -BackgroundColor Black
    Start-Sleep 5

    try {
        $password = Get-Content 'Path...\secret.encrypted'
        return $password
    } catch {
        Write-Host "Error retrieving Room Password: $_" -ForegroundColor Red
        return $null
    }
}

Function SingleRoomInfo {
    param()

    try {
        $displayName = Read-Host 'Enter Room Display Name'
        $alias = Read-Host 'Enter Room Alias/Email Address without @Contoso.com'
        $location = Read-Host 'Enter Room Location'
        $capacity = Read-Host 'Enter Room Capacity'

        $calendarOnly = (Read-Host 'Calendar Only Room? Enter YES or Y for true, anything else for false').ToUpper() -in @('YES', 'Y')

        if ($calendarOnly) {
            $schedulerOnly = $false
            $phoneOnly = $false
            $vc = $false
        } else {
            $schedulerOnly = (Read-Host 'Scheduler Only Room? Enter YES or Y for true, anything else for false').ToUpper() -in @('YES', 'Y')
            $phoneOnly = (Read-Host 'Phone Only Room? Enter YES or Y for true, anything else for false').ToUpper() -in @('YES', 'Y')
            $vc = (Read-Host 'VC Room? Enter YES or Y for true, anything else for false').ToUpper() -in @('YES', 'Y')
        }

        $roomInfo = [PSCustomObject]@{
            DisplayName   = $displayName
            Alias         = $alias
            Location      = $location
            Capacity      = $capacity
            CalendarOnly  = $calendarOnly
            SchedulerOnly = $schedulerOnly
            PhoneOnly     = $phoneOnly
            VC            = $vc
        }

        return $roomInfo
    } catch {
        Write-Host "Error in SingleRoomInfo: $_" -ForegroundColor Red
        return $null
    }
}

Function ImportRoomsInfoCSV {
    Write-Host "Importing Source CSV File for Batch Room Creation" -ForegroundColor Yellow -BackgroundColor Black
    Start-Sleep 5

    try {
        $rooms = Import-Csv -Path 'Path...\RoomsCsv.csv'
        return $rooms
    } catch {
        Write-Host "Error importing Rooms Info CSV: $_" -ForegroundColor Red
        return $null
    }
}

Function CreateRoomMailbox {
    param (
        [string]$displayName,
        [string]$alias,
        [SecureString]$password,
        [int]$capacity,
        [string]$domain = "@contoso.com",
        [string]$usageLocation = "US",
        [int]$bookingWindowInDays = 1080
    )
    Write-Host "Creating Room Mailbox" -ForegroundColor Yellow -BackgroundColor Black
    Start-Sleep -Seconds 5

    try {
        New-Mailbox -DisplayName $displayName -Name $alias -Password $password -ResourceCapacity $capacity -Room
        Start-Sleep -Seconds 30

        Write-Host "Configuring Calendar Processing" -ForegroundColor Yellow -BackgroundColor Black
        Start-Sleep -Seconds 5
        Set-CalendarProcessing -Identity $displayName -BookingWindowInDays $bookingWindowInDays -DeleteSubject $false -DeleteComments $false -AddOrganizerToSubject $false -RemovePrivateProperty $false
        Start-Sleep -Seconds 45

        Write-Host "Setting Usage Location "US"" -ForegroundColor Yellow -BackgroundColor Black
        Start-Sleep -Seconds 5
        $object = Get-MsolUser -UserPrincipalName "$alias$domain"
        $objectID = $object.ObjectId
        Set-MsolUser -ObjectId $objectID -UsageLocation $usageLocation
        Start-Sleep -Seconds 30
    } catch {
        Write-Host "Error creating or configuring mailbox: $_" -ForegroundColor Red
        return $null
    }
}

Function MFAExclusion {
    param ([string]$ObjectID)

    Write-Host "Adding the room to the MFA Exclusion group" -ForegroundColor Yellow -BackgroundColor Black

    try {
        Add-MsolGroupMember -GroupObjectId "ObjectID" -GroupMemberObjectId $ObjectID
    } catch {
        Write-Host "Error adding MFA exclusion: $_" -ForegroundColor Red
    }
}

Function AddLicenseGroup {
    param ([PSCustomObject]$Room, [string]$ObjectID)

    try {
        if ($Room.PhoneOnly -eq $true) {
            Write-Host "Adding to Phone Only License Group" -ForegroundColor Yellow -BackgroundColor Black
            Add-MsolGroupMember -GroupObjectId "ObjectID" -GroupMemberObjectId $ObjectID
        } elseif ($Room.VC -eq $true) {
            Write-Host "Adding to VC License Group" -ForegroundColor Yellow -BackgroundColor Black
            Add-MsolGroupMember -GroupObjectId "ObjectID" -GroupMemberObjectId $ObjectID
        } elseif ($Room.SchedulerOnly -eq $true) {
            Write-Host "Adding to Scheduler Only License Group" -ForegroundColor Yellow -BackgroundColor Black
            Add-MsolGroupMember -GroupObjectId "ObjectID" -GroupMemberObjectId $ObjectID
        } elseif ($Room.CalendarOnly -eq $true) {
            Write-Host "Calendar Only Room" -ForegroundColor Yellow -BackgroundColor Black
        } else {
            Write-Host "No valid group found for this room." -ForegroundColor Red
        }
    } catch {
        Write-Host "Error adding license group: $_" -ForegroundColor Red
    }
}

Function AddDelegations {
    param ([PSCustomObject]$Room, [PSCustomObject]$LocTable, [string]$ObjectID)

    try {
        $Location = $Room.Location
        $Details = $LocTable | Where-Object { ($_.Location -eq $Location) -or ($_.Abbreviation -eq $Location) }

        if ($Details -eq $null) {
            $Details = [PSCustomObject]@{
                Location       = Read-Host 'Enter Room Location Full Name'
                Delegate1      = Read-Host 'Enter Delegate1 UPN'
                Delegate2      = Read-Host 'Enter Delegate2 UPN'
                OutlookRoomlist = Read-Host 'Enter Outlook Roomlist Name'
                State          = Read-Host 'Enter Room State Abbreviation'
                City           = Read-Host 'Enter Room City'
                PostalCode     = Read-Host 'Enter Room Zip/PostalCode'
            }
        }

        $UPN = Get-MsolUser -ObjectId $ObjectID | Select-Object -ExpandProperty UserPrincipalName
        $Location = $Details.Location
        $Delegate1 = $Details.Delegate1
        $Delegate2 = $Details.Delegate2
        $Roomlist = $Details.OutlookRoomlist
        $State = $Details.State
        $City = $Details.City
        $PostalCode = $Details.PostalCode

        Write-Host "Setting delegates for $UPN" -ForegroundColor Yellow -BackgroundColor Black

        if ($Delegate1 -ne $null) {
            add-MailboxPermission -Identity $UPN -User $Delegate1 -AccessRights FullAccess -AutoMapping $true
            Write-Host "Added $Delegate1 as a delegate for $UPN"
            Start-sleep -Seconds 15
        }
        if ($Delegate2 -ne $null) {
            add-MailboxPermission -Identity $UPN -User $Delegate2 -AccessRights FullAccess -AutoMapping $true
            Write-Host "Added $Delegate2 as a delegate for $UPN"
            Start-sleep -Seconds 15
        }

        if ($Roomlist -ne $null) {
            Add-DistributionGroupMember -Identity $Roomlist -Member $UPN
            Write-Host "Added $UPN to the room list $Roomlist"
        }

        Write-Host "Location: $Location"
        Write-Host "State: $State"
        Write-Host "City: $City"
        Write-Host "PostalCode: $PostalCode"
    } catch {
        Write-Host "Error setting delegations: $_" -ForegroundColor Red
    }
}

Function Logs {
    try {
        $Alias = ($NewRoom).Alias
        $UPN = $Alias + "@contoso.com"
        $CreatedRoom = get-msoluser -UserPrincipalName $UPN

        $Logs = [PSCustomObject]@{
            DisplayName = $CreatedRoom.DisplayName
            UPN         = $CreatedRoom.UserPrincipalName
            IsLicensed  = $CreatedRoom.IsLicensed
            CreationDate = (Get-Date).ToString("MM/dd/yyyy")
        }
        return $Logs
    } catch {
        Write-Host "Error generating logs: $_" -ForegroundColor Red
        return $null
    }
}

Function ProcessRoom {
    param (
        [PSCustomObject]$Room,
        [SecureString]$PWD,
        [PSCustomObject]$LocTable,
        [string]$LogPath
    )

    try {
        # Create and configure the room
        CreateAndConfigureRoom -Room $Room -Password $PWD -LocTable $LocTable
        Write-Host "Room $($Room.DisplayName) has been successfully created and configured." -ForegroundColor Green

        # Add the room to the log
        Write-Host "Adding Room to Log" -ForegroundColor Yellow -BackgroundColor Black
        $Log = Logs
        $Log | Export-Csv -Path $LogPath -NoTypeInformation -Append

    } catch {
        Write-Host "Error processing room $($Room.DisplayName): $_" -ForegroundColor Red
    }
}
