# Import Active Directory module
Import-Module ActiveDirectory

# Set the date threshold (30 days ago)
$30DaysAgo = (Get-Date).AddDays(-30)

# Get the current domain name
$domainName = (Get-ADDomain).Name

# Get the current user's desktop path
$desktopPath = "$env:USERPROFILE\Desktop"

# Create the file path for the CSV on the user's desktop
$csvPath = "$desktopPath\$domainName-LastLogonPC.csv"

# Get all enabled computers in Active Directory
$computers = Get-ADComputer -Filter {Enabled -eq $true} -Property LastLogonDate

# Create an array to store the results
$results = @()

# Loop through each computer and check the last logon date
foreach ($computer in $computers) {
    $lastLogon = $computer.LastLogonDate

    # If the computer has a last logon date and it was NOT within the last 30 days
    if ($lastLogon -and $lastLogon -lt $30DaysAgo) {
        # Format the LastLogonDate to show only the date
        $lastLogonDateOnly = $lastLogon.ToString("yyyy-MM-dd")
        
        $results += [pscustomobject]@{
            ComputerName = $computer.Name
            LastLogonDate = $lastLogonDateOnly
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
