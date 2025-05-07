Import-Module ActiveDirectory

$7DaysAgo = (Get-Date).AddDays(-7)
$domainName = (Get-ADDomain).Name
$desktopPath = "$env:USERPROFILE\Desktop"
$csvPath = "$desktopPath\$domainName-Users-Active.csv"
$users = Get-ADUser -Filter * -Property LastLogonDate, Enabled
$results = @()

# Loop through each user and check the last logon date and if the account is enabled
foreach ($user in $users) {
    $lastLogon = $user.LastLogonDate
    $enabled = $user.Enabled

    # If the user has logged in the past week
    if ($lastLogon -and $lastLogon -ge $7DaysAgo -and $enabled) {
        # Format the LastLogonDate to show only the date
        $lastLogonDateOnly = $lastLogon.ToString("yyyy-MM-dd")
        
        $results += [pscustomobject]@{
            UserName = $user.SamAccountName
            LastLogonDate = $lastLogonDateOnly
            IsEnabled = $enabled
        }
    }
}

# Display the results and export them to a CSV file
$results | Sort-Object LastLogonDate -Descending | Format-Table -AutoSize
$results | Export-Csv -Path $csvPath -NoTypeInformation
Write-Host "Report exported to: $csvPath"

# Pause
Read-Host "Press any key to close the window..."