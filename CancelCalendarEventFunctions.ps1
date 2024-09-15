# Cancel Calendar Event Functions

Function GetMeetingOwnerUPN {
    param (
        [string]$Prompt = "Meeting Owner Email Address"
    )
    
    $OwnerUPN = Read-Host $Prompt
    
    # Explicitly return the UPN
    return $OwnerUPN
}

Function GetAttendeeUPN {
    param (
        [string]$Prompt = "Meeting Attendee Email Address"
    )
    
    $AttendeeUPN = Read-Host $Prompt
    
    # Explicitly return the UPN
    return $AttendeeUPN
}

Function GetMeetingSubject {
    param (
        [string]$Prompt = "Meeting Subject (this value has to be exact)"
    )
    
    $Subject = Read-Host $Prompt
    
    # Explicitly return the subject
    return $Subject
}

function GetDPRUserCalendarEvents {
    param([string]$UserId)
    
    # Return the events for the given UserId
    return Get-MgUserEvent -UserId $UserId -All:$true
}

Function TodaysDate {
    return (Get-Date -Format "yyyy-MM-dd_HH-mm")
}

function OwnerCalendarCancellation {
    param([string]$UserId,$EventId)
    
    # Return the result of stopping the event
    return stop-MgUserEvent -UserId $UserId -EventId $EventId
}

function AttendeeMeetingCancellation {
    param([string]$UserId,$EventId)
    
    # Return the result of removing the event
    return Remove-MgUserEvent -UserId $UserId -EventId $EventId
}
