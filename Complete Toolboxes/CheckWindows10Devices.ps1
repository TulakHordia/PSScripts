# Import Active Directory module (if not already imported)
Import-Module ActiveDirectory

$domainName = (Get-ADDomain).Name

# Define the folder path and file path
$folderPath = "C:\Twistech\Script Results"
$filePath = Join-Path -Path $folderPath -ChildPath "$domainName-Windows10Devices.csv"

# Create the folder if it doesn't exist
if (-not (Test-Path -Path $folderPath)) {
    New-Item -Path $folderPath -ItemType Directory | Out-Null
}

# Search for enabled Windows 10 computers in AD
$windows10Computers = Get-ADComputer -Filter {OperatingSystem -like "*Windows 10*" -and Enabled -eq $true} -Property Name,OperatingSystem,OperatingSystemVersion,Enabled

# Display the results
$windows10Computers | Select-Object Name, OperatingSystem, OperatingSystemVersion | Format-Table -AutoSize

# Export the results to CSV
$windows10Computers | Select-Object Name, OperatingSystem, OperatingSystemVersion | Export-Csv -Path $filePath -NoTypeInformation

# Inform the user of the export location
Write-Host "Report exported to: $folderPath"

# Pause the script to prevent the window from closing
Read-Host "Press any key to close the window..."
