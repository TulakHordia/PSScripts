# Connect to Microsoft 365
Connect-ExchangeOnline

# Prompt for the calendar owner's email (User Y)
$calendarOwnerEmail = Read-Host -Prompt "Enter the calendar owner's (User Y) email address"

# Prompt for the delegate or other user's email (User X)
$userEmail = Read-Host -Prompt "Enter the email address of the user (User X) whose permissions you want to check"

# Retrieve specific calendar permissions for User X on User Y's calendar
Write-Output "Checking permissions for $userEmail on $calendarOwnerEmail calendar:"

Get-MailboxFolderPermission -Identity "$calendarOwnerEmail:\Calendar" -User $userEmail | ForEach-Object {
    Write-Output "User: $($_.User) - Access Rights: $($_.AccessRights)"
}

# Disconnect from Microsoft 365
Disconnect-ExchangeOnline

# Pause the script to prevent the window from closing
Read-Host "Press any key to close the window..."