# Requires: ImportExcel module
# Run as Domain Admin for best results
Import-Module ActiveDirectory

# Check and install ImportExcel if missing
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}
Import-Module ImportExcel

$domainName = (Get-ADDomain).Name

# Define folder and file paths
$folderPath = "C:\Twistech\Script Results"
$filePath = Join-Path -Path $folderPath -ChildPath "$domainName-Windows10Devices.xlsx"

# Create the folder if it doesn't exist
if (-not (Test-Path -Path $folderPath)) {
    New-Item -Path $folderPath -ItemType Directory | Out-Null
}

# Query enabled Windows 10 machines
$windows10Computers = Get-ADComputer -Filter {
    OperatingSystem -like "*Windows 10*" -and Enabled -eq $true
} -Property Name, OperatingSystem, OperatingSystemVersion, LastLogonDate

# Prepare results array
$results = foreach ($pc in $windows10Computers) {
    [PSCustomObject]@{
        Name                   = $pc.Name
        OperatingSystem        = $pc.OperatingSystem
        OSVersion              = $pc.OperatingSystemVersion
        LastLogonTimestamp     = $pc.LastLogonDate
        # Placeholder - Requires further setup to get real last logged-on user
        LastLoggedOnUser       = "Unknown"  
    }
}

# Export to XLSX
$results | Export-Excel -Path $filePath -WorksheetName "Windows10Devices" -AutoSize -TableName "Windows10PCs"

# Inform the user
Write-Host "Report exported to: $filePath" -ForegroundColor Green
Read-Host "Press any key to close the window..."
