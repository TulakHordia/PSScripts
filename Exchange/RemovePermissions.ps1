# Import the Exchange Online module
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
Connect-ExchangeOnline -UseDeviceAuthentication

# Prompt for input
$calendarOwner = Read-Host "Enter the calendar owner's email address"
$targetUser = Read-Host "Enter the email address of the user to remove permissions for"

# Remove the user's calendar permissions
try {
    Remove-MailboxFolderPermission -Identity "$calendarOwner:\Calendar" -User $targetUser -Confirm:$false
    Write-Host "Permissions removed for $targetUser from $calendarOwner's calendar." -ForegroundColor Green
} catch {
    Write-Host "Failed to remove permission: $_" -ForegroundColor Red
}

# Disconnect session
Disconnect-ExchangeOnline -Confirm:$false
