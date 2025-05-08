# Import Active Directory module (if not already imported)
Import-Module ActiveDirectory

# Define the folder path and file path
$folderPath = "C:\Twistech\Script Results"
$filePath = Join-Path -Path $folderPath -ChildPath "Windows10_Computers.csv"

# Create the folder if it doesn't exist
if (-not (Test-Path -Path $folderPath)) {
    New-Item -Path $folderPath -ItemType Directory | Out-Null
}

# Search for Windows 10 computers in AD
$windows10Computers = Get-ADComputer -Filter {OperatingSystem -like "*Windows 10*"} -Property Name,OperatingSystem,OperatingSystemVersion

# Display the results
$windows10Computers | Select-Object Name, OperatingSystem, OperatingSystemVersion | Format-Table -AutoSize

# Export the results to CSV
$windows10Computers | Select-Object Name, OperatingSystem, OperatingSystemVersion | Export-Csv -Path $filePath -NoTypeInformation
