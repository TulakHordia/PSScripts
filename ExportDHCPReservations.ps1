# Define DHCP server name (local machine assumed if not specified)
$dhcpServer = "localhost"

# Get all DHCP scopes
$scopes = Get-DhcpServerv4Scope -ComputerName $dhcpServer

# Initialize list for all reservations
$reservationsList = @()

foreach ($scope in $scopes) {
    # Get reservations for each scope
    $reservations = Get-DhcpServerv4Reservation -ComputerName $dhcpServer -ScopeId $scope.ScopeId
    foreach ($reservation in $reservations) {
        # Create a custom object for each reservation
        $reservationsList += [PSCustomObject]@{
            ScopeID       = $scope.ScopeId
            IPAddress     = $reservation.IPAddress
            ClientMAC     = $reservation.ClientId
            Description   = $reservation.Description
            Name          = $reservation.Name
        }
    }
}

# Export the list to CSV
$exportPath = "C:\Twistech\Script Results\DHCP_Reservations.csv"
$reservationsList | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8

Write-Host "Export completed. File saved to $exportPath"