
<# 
.SYNOPSIS
    This script allows canceling a calendar item through MS Graph.
.DESCRIPTION
    Meeting cancellation for terminated/active users via MS Graph.
.AUTHOR
    Amed Aplicano
.DATE
    3/29/2024
#>

# Parameterization
param (
    [Parameter(Mandatory=$true)]
    [string]$OwnerUPN,

    [Parameter(Mandatory=$true)]
    [string]$MeetingSubject
)

# Load Functions
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
Unblock-File -Path "\\abagentqa.dpr.com\F$\SysAdmins\Functions Library\CancelCalendarEventFunctions.ps1" -Confirm:$false
Unblock-File -Path "\\abagentqa.dpr.com\F$\SysAdmins\Functions Library\QAAuthenticationFunctions.ps1" -Confirm:$false

# Load Functions
. "Path...\CancelCalendarEventFunctions.ps1"
. "Path...\AuthenticationFunctions.ps1"

# Modules
Import-module Microsoft.Graph.Calendar
Import-module Microsoft.Graph.Users.Actions

# MS Graph Authentication
MSGraphAuthenticationSchedulerQA

$timestamp = TodaysDate
$Alldomains = Get-MgDomain -All
$domains = $Alldomains.Id

# Loop until a valid OwnerUPN (email address) is provided
while (-not $OwnerUPN -or $OwnerUPN -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
    Write-Host "Invalid OwnerUPN format. Please provide a valid email address."
    $OwnerUPN = Read-Host "Enter the Meeting Owner Email Address"
}

# Loop until a valid MeetingSubject (non-empty) is provided
while ([string]::IsNullOrWhiteSpace($MeetingSubject)) {
    Write-Host "MeetingSubject cannot be empty. Please provide a valid meeting subject."
    $MeetingSubject = Read-Host "Enter the Meeting Subject (this value has to be exact)"
}

# Fetch events
$events = GetDPRUserCalendarEvents -UserId $OwnerUPN

# Filter events using keyword matching
$FilteredEvents = $events | Where-Object {$_.Subject -like "*$MeetingSubject*"}

if ($FilteredEvents.Count -gt 1) {
    Write-Host "Multiple events found with matching keywords. Please select the event you wish to cancel."

    # Display each event with a number
    for ($i = 0; $i -lt $FilteredEvents.Count; $i++) {
        $event = $FilteredEvents[$i]
        $startDate = Get-Date $event.Start.DateTime -Format "MM/dd/yyyy"
        Write-Host "$i. Subject: $($event.Subject), Start Time: $startDate, Organizer: $($event.Organizer.EmailAddress.Address)"
    }

    # Prompt the user to choose the event to cancel
    $selectedIndex = Read-Host "Enter the number of the meeting you wish to cancel"

    # Validate the user input
    while (-not [int]::TryParse($selectedIndex, [ref]$null) -or $selectedIndex -lt 0 -or $selectedIndex -ge $FilteredEvents.Count) {
        Write-Host "Invalid selection. Please enter a valid number."
        $selectedIndex = Read-Host "Enter the number of the meeting you wish to cancel"
    }

    # Store the selected event in a variable
    $selectedEvent = $FilteredEvents[$selectedIndex]
    $selectedStartDate = Get-Date $selectedEvent.Start.DateTime -Format "MM/dd/yyyy"
    Write-Host "You have selected: Subject: $($selectedEvent.Subject), Start Time: $selectedStartDate"
} elseif ($FilteredEvents.Count -eq 1) {
    # Only one event found, proceed with it
    $selectedEvent = $FilteredEvents[0]
    $selectedStartDate = Get-Date $selectedEvent.Start.DateTime -Format "MM/dd/yyyy"
    Write-Host "Only one event found: Subject: $($selectedEvent.Subject), Start Time: $selectedStartDate"
} else {
    Write-Host "No matching events found."
}

# Now $selectedEvent contains the event to cancel
# You can use it in the cancellation function
try {
    OwnerCalendarCancellation -UserId $OwnerUPN -EventId $selectedEvent.Id
    Write-Host "Successfully canceled the selected event."
} catch {
    Write-Error "Failed to cancel the event: $($_.Exception.Message)"
}

# Disconnect from Graph
try {
    Disconnect-Graph
    Write-Host "Successfully disconnected from Microsoft Graph."
} catch {
    Write-Error "Failed to disconnect from Microsoft Graph: $($_.Exception.Message)"
}
