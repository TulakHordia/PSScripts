# Import Exchange Online module if not already loaded
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
}
Import-Module ExchangeOnlineManagement -Force

# Global Variables
$domainName = (Get-ADDomain).Name
$folderPath = "C:\Twistech\Script Results"
$csvPath = "$folderPath\$domainName-Exchange.csv"

function Show-Menu {
    Clear-Host
    Write-Host "===================================" -ForegroundColor DarkCyan
    Write-Host "         Exchange Toolbox           " -ForegroundColor Green
    Write-Host "===================================" -ForegroundColor DarkCyan
    Write-Host "1. Connect to Exchange Online"
    Write-Host " ------ Calendar ------"
    Write-Host "2. Check Calendar Permissions for a User"
    Write-Host "3. Give Calendar Permission - EDITOR"
    Write-Host "4. Give Calendar Permission - REVIEWER"
    Write-Host "5. Remove Calendar Permissions"
    Write-Host " ------ Mailbox ------"
    Write-Host "6. Export Mailbox Size"
    Write-Host " ------ Change Exchange Variables ------"
    Write-Host "7. Add an alias to all mailboxes"
    Write-Host "Q. Quit"
    Write-Host ""
}

do {
    Show-Menu
    $choice = Read-Host "Select an option"
    switch ($choice) {
        '1' { Connect-ExchangeOnlineSession }
        '2' { Check-CalendarPermissions }
        '3' { Give-EditorPermission }
        '4' { Give-ReviewerPermission }
        '5' { Remove-CalendarPermissions }
        '6' { Export-MailboxSize}
        '7' { Add-AliasToAllMailboxes}
        'Q' { Write-Host "Exiting script." -ForegroundColor Cyan }
        default { Write-Host "Invalid selection. Please try again." -ForegroundColor Red }
    }
    if ($choice -ne 'Q') {
        Write-Host "`nPress Enter to return to the menu..."
        [void][System.Console]::ReadLine()
    }
} while ($choice -ne 'Q')

function Connect-ExchangeOnlineSession {
    try {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
        Connect-ExchangeOnline
        Write-Host "Connected successfully!" -ForegroundColor Green
    } catch {
        Write-Host "Failed to connect to Exchange Online." -ForegroundColor Red
    }
}

function Give-EditorPermission {
    $calendarOwner = Read-Host "Enter the email address of the calendar owner"
    $grantPermissions = $true

    while ($grantPermissions) {
        $recipient = Read-Host "Enter the email address of the person to grant Editor permissions to"

        Add-MailboxFolderPermission -Identity "$($calendarOwner):\Calendar" -User $recipient -AccessRights Editor
        Add-MailboxFolderPermission -Identity "$($calendarOwner):\לוח שנה" -User $recipient -AccessRights Editor

        $WarningActionPreference = "Continue"

        $response = Read-Host "Do you want to grant permissions to another user? (Y/N)"
        if ($response -ne "Y") {
            $grantPermissions = $false
        }
    }
}

function Give-ReviewerPermission {
    $calendarOwner = Read-Host "Enter the email address of the calendar owner"
    $grantPermissions = $true

    while ($grantPermissions) {
        $recipient = Read-Host "Enter the email address of the person to grant Editor permissions to"

        Add-MailboxFolderPermission -Identity "$($calendarOwner):\Calendar" -User $recipient -AccessRights Reviewer
        Add-MailboxFolderPermission -Identity "$($calendarOwner):\לוח שנה" -User $recipient -AccessRights Reviewer

        $WarningActionPreference = "Continue"

        $response = Read-Host "Do you want to grant permissions to another user? (Y/N)"
        if ($response -ne "Y") {
            $grantPermissions = $false
        }
    }
}

function Check-CalendarPermissions {
    $UserEmail = Read-Host "Enter the email address to check calendar permissions for"
    
    # Get all mailboxes
    $Mailboxes = Get-Mailbox -ResultSize Unlimited
    $Results = @()

    foreach ($Mailbox in $Mailboxes) {
        $CalendarPermissions = Get-MailboxFolderPermission -Identity "$($Mailbox.PrimarySmtpAddress):\Calendar" -ErrorAction SilentlyContinue
        foreach ($Permission in $CalendarPermissions) {
            if ($Permission.User -like "*$UserEmail*") {
                $Results += [PSCustomObject]@{
                    Mailbox      = $Mailbox.PrimarySmtpAddress
                    AccessRights = $Permission.AccessRights -join ", "
                    User         = $Permission.User
                }
            }
        }
    }

    if ($Results.Count -gt 0) {
        $Results | Format-Table -AutoSize
        $export = Read-Host "Export to CSV? (Y/N)"
        if ($export -eq "Y") {
            $csvPath = "C:\Twistech\Script Results\CalendarPermissions_$($UserEmail.Replace('@','_')).csv"
            $Results | Export-Csv -Path $csvPath -NoTypeInformation
            Write-Host "Results exported to $csvPath" -ForegroundColor Green
        }
    } else {
        Write-Host "No calendar permissions found for $UserEmail" -ForegroundColor Yellow
    }
}

function Remove-CalendarPermissions {
    $calendarOwner = Read-Host "Enter the email address of the calendar owner"
    $grantPermissions = $true

    while ($grantPermissions) {
        $userToRemove = Read-Host "Enter the email address of the user whose permissions you want to revoke"

        Remove-MailboxFolderPermission -Identity "$($calendarOwner):\Calendar" -User $userToRemove
        Remove-MailboxFolderPermission-MailboxFolderPermission -Identity "$($calendarOwner):\לוח שנה" -User $userToRemove

        $WarningActionPreference = "Continue"

        $response = Read-Host "Do you want to revoke permissions from another user? (Y/N)"
        if ($response -ne "Y") {
            $grantPermissions = $false
        }
    }
}

function Export-MailboxSize {
$Mailboxes = Get-Mailbox -ResultSize Unlimited

$MailboxData = $Mailboxes | ForEach-Object {
    $Stats = Get-MailboxStatistics $_.PrimarySmtpAddress
    $ArchiveStats = Get-MailboxStatistics $_.PrimarySmtpAddress -Archive
    [PSCustomObject]@{
        DisplayName      = $_.DisplayName
        PrimarySMTP      = $_.PrimarySmtpAddress
        MailboxType      = $_.RecipientTypeDetails
        TotalItemSize    = $Stats.TotalItemSize.ToString()
        ItemCount        = $Stats.ItemCount
        ArchiveEnabled   = if ($_.ArchiveStatus -eq "Active") {"Yes"} else {"No"}
        ArchiveSize      = if ($ArchiveStats) { $ArchiveStats.TotalItemSize.ToString() } else {"N/A"}
    }
}

$MailboxData | Export-Csv -Path "$folderPath\$domainName-MailboxReport.csv" -NoTypeInformation

}

function Add-AliasToAllMailboxes {
    Connect-MsolService

    $domain = Read-Host "Enter the domain to use for the alias (e.g. example.com)"

    Write-Host "`nChoose an option:"
    Write-Host "1. Add alias to a single user"
    Write-Host "2. Add alias to all licensed users"
    $subChoice = Read-Host "Enter your choice"

    switch ($subChoice) {
        '1' {
            $userUPN = Read-Host "Enter the UPN (email address) of the user"
            $alias = $userUPN -replace '@.*$', "@$domain"
            try {
                Set-Mailbox -Identity $userUPN -EmailAddresses @{add = $alias}
                Write-Host "Alias $alias added to $userUPN" -ForegroundColor Green
            } catch {
                Write-Host "Failed to add alias: $_" -ForegroundColor Red
            }
        }
        '2' {
            $users = Get-MsolUser -All | Where-Object { $_.isLicensed -eq $true }
            foreach ($user in $users) {
                $primarySMTP = $user.UserPrincipalName
                $alias = $primarySMTP -replace '@.*$', "@$domain"
                try {
                    Set-Mailbox -Identity $primarySMTP -EmailAddresses @{add = $alias}
                    Write-Host "Alias $alias added to $primarySMTP" -ForegroundColor Green
                } catch {
                    Write-Host "Failed to add alias to $primarySMTP : $_" -ForegroundColor Red
                }
            }
        }
        default {
            Write-Host "Invalid option selected." -ForegroundColor Red
        }
    }

    Read-Host -Prompt "`nPress Enter to return to the menu..."
}
