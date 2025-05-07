# Import Active Directory module
Import-Module ActiveDirectory

# Get the current domain name
$domainName = (Get-ADDomain).Name

# Get the current user's desktop path
$desktopPath = "$env:USERPROFILE\Desktop"

# Create the file path for the CSV on the user's desktop
$csvPath = "$desktopPath\$domainName-DisabledUsers.csv"

# Get all disabled users in Active Directory with necessary properties
$users = Get-ADUser -Filter {Enabled -eq $false} -Property SamAccountName | Export-Csv -path $csvPath -NoTypeInformation

# Inform the user of the export location
Write-Host "Report exported to: $csvPath"

# Pause the script to prevent the window from closing
Read-Host "Press any key to close the window..."
