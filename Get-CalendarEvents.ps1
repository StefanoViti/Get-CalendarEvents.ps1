<#
.SYNOPSIS
Get-CalendarEvents.ps1 - This script retrieves the calendar events of a user or a list of users.

.DESCRIPTION 
This script retrieves the calendar events of a user or a list of users. You can save a file named UsersList.csv
in your the same script folder, containing a single column with the User Principal Name of the users you want the calendar 
and "UserPrincipalName" as header. Otherwise, the script will ask you to provide the UPN of the single user

.OUTPUTS
A final csv report will be provided at .\UsersCalendarEvents.csv

.REQUIREMENTS
Before running the script, you have to create an app registration with a secret in your Microsoft Entra tenant.
The application must have the following application permissions (with admin consent granted):
- Calendars.Read
- OnlineMeetings.Read.All
- User.Read.All

.PARAMETER AppId
Insert the application ID (clientID) of your app registration

.PARAMETER TenantId
Insert the ID of your tenant

.PARAMETER ClientSecret
Insert the Secret created for your application

.EXAMPLE
.\Get-CalendarEvents.ps1 -AppId "01101d11-01x1-1011-1011-101101z01011" -TenantID "1a1b1c11-1d11-f111-1h1y1-1u11919111411" -ClientSecret "xxxxxx"

.NOTES
Written by: Stefano Viti - stefano.viti1995@gmail.com
Follow me at https://www.linkedin.com/in/stefano-viti/

#>

param(
    [Parameter(Mandatory=$True)]
    [string]$AppId,

	[Parameter(Mandatory=$True)]
    [switch]$TenantID,

    [Parameter(Mandatory=$True)]
    [switch]$ClientSecret
)

clear

#Script to Retrieve Calendar Events for users

#Get an access token

$ReqToken = @{
      Grant_Type = "Client_Credentials"
      client_id = $AppId
      client_secret = $ClientSecret
      Scope = "https://graph.microsoft.com/.default"
}

$Uri = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"
$Resp = Invoke-RestMethod -Uri $Uri -Method Post -Body $ReqToken -ContentType "application/x-www-form-urlencoded"

$Token = $Resp.access_token

#Get the Calendar Events of a speficic user or upload a csv file for more users

try{
    $Users = Import-csv -path ".\UsersList.csv" -ErrorAction Stop
    }
catch{
    $UPN = Read-Host -Prompt "Insert the UPN of the user of which you want to get the calendar"   
    }

$Results = @()

if($Users -ne $null){

    $Path = ".\UsersCalendarEvents.csv"

    Foreach($User in $Users){
        $UPN = $User.UserPrincipalName
        $Api = "https://graph.microsoft.com/v1.0/users/$UPN/events"
        Write-Host $api -ForegroundColor Yellow
        $Headers = @{
            Authorization = "Bearer $Token"
        }

        $Events = Invoke-RestMethod -Headers $Headers -Uri $Api -Method Get -ContentType application/json

        #Export the Results

        $ResultsEvents = @()

        foreach($Event in $Events.value){
            $Attendees = $Event.attendees.emailaddress
            $AttendeesList = @()
            Foreach ($Attende in $Attendees){
                $AttendeMailAddress = $Attende.Address
                $AttendeesList = $AttendeesList + $AttendeMailAddress
            }
            $AttendeesList = $AttendeesList -join "|"
            $hash = [ordered]@{
                User = $UPN
                Id = $Event.id
                IsTeamsMeeting = $Event.isOnlineMeeting
                Subject = $Event.subject
                StartTime = [datetime]$Event.start.datetime
                EndTime = [datetime]$Event.end.datetime
                OrganizerName = $Event.organizer.emailaddress.name
                OrganizerMailAddress = $Event.organizer.emailaddress.address
                Attendees = $AttendeesList
                AttendeesCount = ($Event.attendees).count
            }
            $Item = New-Object PSObject -Property $hash
            $ResultsEvents = $ResultsEvents + $Item
        }

        $Results = $Results + $ResultsEvents

    }

}

else{
    $Path = ".\CalendarEvents_$UPN.csv"
     
    $Api = "https://graph.microsoft.com/v1.0/users/$UPN/events"
    #https://graph.microsoft.com/v1.0/users/$userId/calendar/events?`$filter=start/dateTime ge '$($startDate)' and start/dateTime lt '$($endDate)'
    Write-Host $api -ForegroundColor Yellow
    $Headers = @{
          Authorization = "Bearer $Token"
    }
    
    $Events = Invoke-RestMethod -Headers $Headers -Uri $Api -Method Get -ContentType application/json
    
    #Export the Results
    
    $ResultsEvents = @()
    
    foreach($Event in $Events.value){
        $Attendees = $Event.attendees.emailaddress
        $AttendeesList = @()
        Foreach ($Attende in $Attendees){
            $AttendeMailAddress = $Attende.Address
            $AttendeesList = $AttendeesList + $AttendeMailAddress
        }
        $AttendeesList = $AttendeesList -join "|"
        $hash = [ordered]@{
            User = $UPN
            Id = $Event.id
            IsTeamsMeeting = $Event.isOnlineMeeting
            Subject = $Event.subject
            StartTime = [datetime]$Event.start.datetime
            EndTime = [datetime]$Event.end.datetime
            OrganizerName = $Event.organizer.emailaddress.name
            OrganizerMailAddress = $Event.organizer.emailaddress.address
            Attendees = $AttendeesList
            AttendeesCount = ($Event.attendees).count
        }
        $Item = New-Object PSObject -Property $hash
        $ResultsEvents = $ResultsEvents + $Item
    }
    
    $Results = $Results + $ResultsEvents
}

$Results | Export-csv -path $Path