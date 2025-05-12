# Registry path for Outlook profiles
$regPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"

# Check if the Outlook profiles registry path exists
if (Test-Path $regPath) {
    # Get all profile names
    $profiles = Get-ChildItem -Path $regPath | Select-Object -ExpandProperty PSChildName

    # Find profiles that match the pattern (ends with " - Google Workspace")
    $matchingProfiles = $profiles | Where-Object { $_ -like "* - Google Workspace *" }

    if ($matchingProfiles) {
        foreach ($profile in $matchingProfiles) {
            Write-Output "Removing Outlook Profile: $profile"
            Remove-Item -Path "$regPath\$profile" -Recurse -Force
        }
    } else {
        Write-Output "No Google Workspace profiles found."
    }
} else {
    Write-Output "Outlook Profiles registry path not found."
}

# Pause at the end
Write-Output "Press Enter to exit..."
Read-Host
