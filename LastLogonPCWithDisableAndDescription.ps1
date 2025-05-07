# Import Active Directory module
Import-Module ActiveDirectory

# Set the date threshold (30 days ago)
$30DaysAgo = (Get-Date).AddDays(-30)

# Get the current domain name
$domainName = (Get-ADDomain).Name

# Get the current user's desktop path
$desktopPath = "C:\Twistech\Script Results"

# Create the file path for the CSV on the user's desktop
$csvPath = "$desktopPath\$domainName-LastLogonPC.csv"

# Get all enabled computers in Active Directory
$computers = Get-ADComputer -Filter {Enabled -eq $true} -Property LastLogonDate, Description

# Create an array to store the results
$results = @()

# Loop through each computer and check the last logon date
foreach ($computer in $computers) {
    $lastLogon = $computer.LastLogonDate

    # If the computer has a last logon date and it was NOT within the last 30 days
    if ($lastLogon -and $lastLogon -lt $30DaysAgo) {
        # Format the LastLogonDate to show only the date
        $lastLogonDateOnly = $lastLogon.ToString("yyyy-MM-dd")
        
        # Add the computer to the results array
        $results += [pscustomobject]@{
            ComputerName = $computer.Name
            LastLogonDate = $lastLogonDateOnly
        }

        # Disable the computer account
        Disable-ADAccount -Identity $computer.DistinguishedName

        # Update the description
        Set-ADComputer -Identity $computer.DistinguishedName -Description "Disabled by Beni - 07/05/2025"
    }
}

# Display the results
$results | Sort-Object LastLogonDate -Descending | Format-Table -AutoSize

# Export to CSV
$results | Export-Csv -Path $csvPath -NoTypeInformation

# Inform the user
Write-Host "Report exported to: $csvPath"
Write-Host "Accounts disabled and descriptions updated."

# Pause the script
Read-Host "Press any key to close the window..."
