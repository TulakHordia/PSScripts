$UserEmail = "input_user_email"  # Change this to the user you want to check

# Get all mailboxes
$Mailboxes = Get-Mailbox -ResultSize Unlimited

# Initialize results array
$Results = @()

# Loop through each mailbox and check calendar permissions
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

# Output results to CSV
$Results | Format-Table -AutoSize
