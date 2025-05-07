# Import Active Directory module
Import-Module ActiveDirectory

# Set the date threshold (30 days ago)
$30DaysAgo = (Get-Date).AddDays(-30)

# Get the current domain name
$domainName = (Get-ADDomain).Name

# Get the current user's desktop path
$desktopPath = "$env:USERPROFILE\Desktop"

# Create the file path for the CSV on the user's desktop
$csvPath = "$desktopPath\$domainName-Users-LastLogon.csv"

# Get all users in Active Directory with necessary properties
$users = Get-ADUser -Filter * -Property LastLogonDate, Enabled

# Create an array to store the results
$results = @()

# Loop through each user and check the last logon date and if the account is enabled
foreach ($user in $users) {
    $lastLogon = $user.LastLogonDate
    $enabled = $user.Enabled

    # If the user has a last logon date older than 30 days and the account is enabled
    if ($lastLogon -and $lastLogon -lt $30DaysAgo -and $enabled) {
        # Format the LastLogonDate to show only the date
        $lastLogonDateOnly = $lastLogon.ToString("yyyy-MM-dd")
        
        $results += [pscustomobject]@{
            UserName = $user.SamAccountName
            LastLogonDate = $lastLogonDateOnly
            IsEnabled = $enabled
        }
    }
}

# Display the results or export them to a CSV file
$results | Sort-Object LastLogonDate -Descending | Format-Table -AutoSize

# Export to CSV with the domain name in the file path
$results | Export-Csv -Path $csvPath -NoTypeInformation

# Inform the user of the export location
Write-Host "Report exported to: $csvPath"

# Pause the script to prevent the window from closing
Read-Host "Press any key to close the window..."