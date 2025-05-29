# Load Exchange Online module
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module ExchangeOnlineManagement -Force -Scope CurrentUser
}
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
try {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Connect-ExchangeOnline
    Write-Host "Connected successfully." -ForegroundColor Green
} catch {
    Write-Host "Failed to connect to Exchange Online: $_" -ForegroundColor Red
    exit
}

do {
    # Prompt for user input
    $calendarOwner = Read-Host "Enter the calendar owner's email address"
    $delegateUser = Read-Host "Enter the email address of the user to grant access"

    # Choose permission level
    do {
        $accessRight = Read-Host "Enter the access right (Editor or Reviewer)"
    } while ($accessRight -notin @("Editor", "Reviewer"))

    # Grant permissions
    try {
        Add-MailboxFolderPermission -Identity "$($calendarOwner):\Calendar" -User $delegateUser -AccessRights $accessRight
        Write-Host "$accessRight permissions granted to $delegateUser on $calendarOwner's calendar." -ForegroundColor Green
    } catch {
        Write-Host "Failed to grant permission: $_" -ForegroundColor Red
    }

    # Ask if user wants to do another
    $continue = Read-Host "Do you want to grant more permissions? (Y/N)"
} while ($continue -match "^[Yy]")

# Disconnect session
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Disconnected from Exchange Online." -ForegroundColor Cyan