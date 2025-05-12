# Get all Windows devices enrolled in Intune
$devices = Get-MgDeviceManagementManagedDevice -All | Where-Object {
    $_.OperatingSystem -eq "Windows" -and $_.EncryptionState -eq 1
}

# Build the filtered report
$bitlockerDevices = $devices | ForEach-Object {
    [PSCustomObject]@{
        DeviceName        = $_.DeviceName
        UserPrincipalName = $_.UserPrincipalName
        OSVersion         = $_.OsVersion
        BitLockerStatus   = "Enabled"
        LastSyncDate      = $_.LastSyncDateTime
        ComplianceState   = $_.ComplianceState
    }
}

# Output or export
$bitlockerDevices | Format-Table -AutoSize
$bitlockerDevices | Export-Csv -Path "S:\Tech Team\Benjamin\Zoozpower\DeviceInventoryReport.csv" -NoTypeInformation