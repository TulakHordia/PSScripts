# Define the process name to search for
$ProcessName = "SentinelAgent.exe" # Replace with the actual process name

# Get a list of all computers in the domain
$Computers = Get-ADComputer -Filter * | Select-Object -ExpandProperty Name

# Output file
$OutputFile = "C:\SentinelOne_Process_Results.txt"

foreach ($Computer in $Computers) {
    try {
        Write-Host "Checking $Computer..."
        $Processes = Get-Process -ComputerName $Computer | Where-Object { $_.Name -eq $ProcessName }
        if ($Processes) {
            Add-Content -Path $OutputFile -Value "$Computer is running $ProcessName"
        } else {
            Add-Content -Path $OutputFile -Value "$Computer is not running $ProcessName"
        }
    } catch {
        Add-Content -Path $OutputFile -Value "$Computer: Unable to connect or query processes"
    }
}


# Inform the user of the export location
Write-Host "Report exported to: $OutputFile"

# Pause the script to prevent the window from closing
Read-Host "Press any key to close the window..."