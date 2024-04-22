# Get-CalendarEvents.ps1
This PowerShell script retrieves the Outlook calendar events of a single user or a list of users. Please read the synopsis at the beginning of the script for more details and information about the inputs and the outputs.

Before running the script, you have to create an app registration with a secret in your Microsoft Entra tenant.
The application must have the following application permissions (with admin consent granted):
- Calendars.Read
- OnlineMeetings.Read.All
- User.Read.All

Written by: Stefano Viti - stefano.viti1995@gmail.com
Follow me at https://www.linkedin.com/in/stefano-viti/
