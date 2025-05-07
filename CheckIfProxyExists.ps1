Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All"
$proxyAddress = "input_smtp"  # Replace with the actual email
Get-MgUser -All | Where-Object { $_.ProxyAddresses -contains "SMTP:$proxyAddress" } | Select-Object DisplayName, UserPrincipalName, ProxyAddresses
