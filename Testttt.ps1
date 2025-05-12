$userUPN = "Itay.Rosenfeld1@TCSM.onmicrosoft.com"  # Replace with actual UPN
Get-MgUser -UserId $userUPN | Select-Object ProxyAddresses

$userUPN2 = "itay.rosenfeld@techsomed.com"  # Replace with actual UPN
Get-MgUser -UserId $userUPN2 | Select-Object ProxyAddresses
