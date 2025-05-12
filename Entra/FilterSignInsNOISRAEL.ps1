# Import Microsoft Graph Module
Import-Module Microsoft.Graph

# Connect to Microsoft Graph (you will be prompted to sign in)
Connect-MgGraph -Scopes "AuditLog.Read.All"

# Define the time range (last 7 days)
$startDateTime = (Get-Date).AddDays(-7).ToString("o")
$endDateTime = (Get-Date).ToString("o")

# Retrieve sign-in logs for the last 7 days
$signInLogs = Get-MgAuditLogSignIn -Filter "createdDateTime ge $startDateTime and status/errorCode eq 0" -All

# Filter logs for successful sign-ins from countries other than Israel
$filteredSignInLogs = $signInLogs | Where-Object {
    $_.location.countryOrRegion -ne "Israel"
}

# Export the results to CSV
$filteredSignInLogs | Select-Object userDisplayName, userPrincipalName, createdDateTime, location |
    Export-Csv -Path "FilteredSignInLogs.csv" -NoTypeInformation -Encoding UTF8

Write-Output "Export completed. Results saved to 'FilteredSignInLogs.csv'."

# Disconnect from Microsoft Graph
Disconnect-MgGraph

# Pause at the end
Read-Host -Prompt "Press Enter to exit"
