Connect-MsolService -Credential $UserCredential
Connect-ExchangeOnline -Credential $UserCredential

$users = Get-MsolUser -All | Where-Object { $_.isLicensed -eq $true }
$domain = "input_domain_here"

foreach ($user in $users) {
    $primarySMTP = $user.UserPrincipalName
    $alias = $primarySMTP -replace '@.*$', "@$domain"
    
    # Add alias
    Set-Mailbox -Identity $primarySMTP -EmailAddresses @{add=$alias}
}

# Pause the script at the end
Read-Host -Prompt "Press Enter to exit..."
