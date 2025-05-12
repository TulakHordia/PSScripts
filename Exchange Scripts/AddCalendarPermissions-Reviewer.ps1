# Check if the Exchange Online PowerShell module is installed
if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
    # If not installed, prompt user to install it
    Write-Host "The Exchange Online PowerShell module is not installed. Would you like to install it now? (Y/N)"
    $installResponse = Read-Host
    if ($installResponse -eq "Y") {
        # Install the Exchange Online PowerShell module
        Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
    }
    else {
        Write-Host "Exiting script. Please install the Exchange Online PowerShell module and run the script again."
        exit
    }
}

# Import the Exchange Online PowerShell module
Import-Module ExchangeOnlineManagement -Force

# Prompt for credentials
$creds = Get-Credential

# Connect to Exchange Online
Connect-ExchangeOnline -Credential $creds

# Prompt for the user whose calendar you want to grant permissions to (Calendar Owner)
$calendarOwner = Read-Host "Enter the email address of the user whose calendar you want to share: (e.g., user@example.com):"

# Initialize variables
$grantPermissions = $true

# Loop to grant permissions
while ($grantPermissions) {
    # Prompt for the user to receive permissions (Recipient)
    $recipient = Read-Host "Enter the email address of the person who will receive the permissions (e.g., user@example.com):"

    # Suppress warnings
    $WarningActionPreference = "SilentlyContinue"
    
    # Try to add permissions to the calendar
    Add-MailboxFolderPermission -Identity "$($calendarOwner):\Calendar" -User $recipient -AccessRights Reviewer
    Add-MailboxFolderPermission -Identity "$($calendarOwner):\לוח שנה" -User $recipient -AccessRights Reviewer

    # Reset warning preference
    $WarningActionPreference = "Continue"

    # Prompt to grant permissions to another calendar
    $response = Read-Host "Do you want to grant permissions to another person? (Y/N)"
    if ($response -ne "Y") {
        $grantPermissions = $false
    }
}

