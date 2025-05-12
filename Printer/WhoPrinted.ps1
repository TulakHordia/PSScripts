# Define the path for the operational log
$logName = "Microsoft-Windows-PrintService/Operational"

# Define the output CSV file
$outputCsv = "PrintEventLogs.csv"

# Check if the log is enabled
if (-not (Get-WinEvent -ListLog $logName -ErrorAction SilentlyContinue)) {
    Write-Host "The '$logName' log is not enabled or accessible. Please ensure it's enabled in Event Viewer."
    return
}

# Get Event ID 307 from the log
Write-Host "Retrieving Event ID 307 logs..."
$events = Get-WinEvent -LogName $logName -FilterXPath "*[System[(EventID=307)]]" -ErrorAction SilentlyContinue

if ($events.Count -eq 0) {
    Write-Host "No Event ID 307 logs found."
    return
}

# Process events and extract relevant information
$logData = $events | ForEach-Object {
    # Parse XML to extract additional details
    $eventXml = [xml]$_.ToXml()

    # Extract the actual username from the event
    $username = $eventXml.Event.UserData.DocumentPrinted.Param3

    # Extract the printer from the event
    $printer = $eventXml.Event.UserData.DocumentPrinted.Param5

    # Handle cases where the username might not be available
    if ([string]::IsNullOrWhiteSpace($username)) {
        $username = "Unknown"
    }
	
	# Format the 'Logged' time to only include the date
    $loggedDate = $_.TimeCreated.ToString('yyyy-MM-dd')

    # Create a custom object to hold the relevant information
    [PSCustomObject]@{
        User     = $username
	Printer  = $printer
        Logged   = $_.TimeCreated
    }
}

# Export to CSV
if ($logData.Count -gt 0) {
    $logData | Export-Csv -Path $outputCsv -NoTypeInformation -Encoding UTF8
    Write-Host "Logs exported to '$outputCsv'."
} else {
    Write-Host "No data to export."
}

# Pause the script to prevent the window from closing
Read-Host "Press any key to close the window..."
