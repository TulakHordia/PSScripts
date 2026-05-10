# Import Active Directory module
Import-Module ActiveDirectory

# Define export path
$exportPath = "C:\Informat\Script Results\DisabledUsers.csv"

# Create export folder if it doesn't exist
$folder = Split-Path -Path $exportPath
if (-not (Test-Path -Path $folder)) {
    New-Item -Path $folder -ItemType Directory | Out-Null
}

# Get all disabled users
$disabledUsers = Get-ADUser -Filter {Enabled -eq $false} -Properties Name, SamAccountName, UserPrincipalName, Enabled, Description, LastLogonDate

# Export to CSV
$disabledUsers |
    Select-Object Name, SamAccountName, UserPrincipalName, Enabled, Description, LastLogonDate |
    Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8

Write-Host "Export completed. File saved to $exportPath" -ForegroundColor Green
