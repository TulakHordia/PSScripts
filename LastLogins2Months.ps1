# Define the date threshold for logon events (2 months ago)
$twoMonthsAgo = (Get-Date).AddMonths(-2)

# Define the output CSV file path (saving to the Desktop)
$outputFile = "$env:USERPROFILE\Desktop\CurrentMachineLogins.csv"

# Get the logon events from the Security log
$logonEvents = Get-WinEvent -LogName Security -FilterXPath "*[System[TimeCreated[@SystemTime >= '$($twoMonthsAgo.ToUniversalTime().ToString("o"))']]]" `
    | Where-Object { $_.Id -eq 4624 }  # Event ID 4624 corresponds to a successful logon

# Extract relevant details and filter out SYSTEM/Service accounts
$filteredEvents = $logonEvents | ForEach-Object {
    $userName = ($_.Properties[5].Value -split '\\')[-1] # Extract username
    $logonType = $_.Properties[8].Value                 # Extract logon type

    if ($userName -notlike "SYSTEM" -and $userName -notlike "LOCAL SERVICE" -and $userName -notlike "NETWORK SERVICE" -and $userName -notlike "ttadmin" -and $userName -notlike "RH-VALIDATION$" -and $userName -notlike "Administrator" -and $userName -notlike "RH-LYOSERVER$") {
        [PSCustomObject]@{
            TimeCreated = $_.TimeCreated
            UserName    = $userName
            LogonType   = $logonType
        }
    }
}

# Export the filtered results to CSV
$filteredEvents | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8

Write-Output "Non-system user login data exported to $outputFile"

# Pause at the end of the script
Read-Host "Press Enter to exit"
