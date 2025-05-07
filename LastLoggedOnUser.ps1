# Define the parameters for the script
$LogFilePath = "ADUsersLastLogonWithMachines.csv"
$DomainController = "localhost" # Change if querying a remote DC
$LogName = "Security"
$EventID = 4624 # Logon event ID

# Create an array to store results
$results = @()

# Get logon events from the Security log
Write-Host "Querying the Security log on $DomainController for Event ID $EventID. This may take some time..."
$events = Get-WinEvent -LogName $LogName -FilterXPath "*[System[(EventID=$EventID)]]" -ComputerName $DomainController

foreach ($event in $events) {
    # Parse the event details
    $xml = [xml]$event.ToXml()
    $username = $xml.Event.EventData.Data | Where-Object { $_.Name -eq "TargetUserName" } | Select-Object -ExpandProperty '#text'
    $computer = $xml.Event.EventData.Data | Where-Object { $_.Name -eq "WorkstationName" } | Select-Object -ExpandProperty '#text'
    $logonTime = $event.TimeCreated

    if ($username -and $computer) {
        # Add the details to the results
        $results += [PSCustomObject]@{
            UserName   = $username
            Computer   = $computer
            LogonTime  = $logonTime
        }
    }
}

# Export the results to a CSV file
$results | Sort-Object UserName, LogonTime | Export-Csv -Path $LogFilePath -NoTypeInformation -Encoding UTF8

Write-Host "The script has completed. Results are saved in '$LogFilePath'."

Read-Host "Pausing"