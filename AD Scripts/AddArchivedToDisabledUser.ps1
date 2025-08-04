# Import Active Directory module
Import-Module ActiveDirectory

# Get all disabled users
$disabledUsers = Get-ADUser -Filter 'Enabled -eq $false' -Properties DisplayName

foreach ($user in $disabledUsers) {
    # Check if DisplayName is already archived
    if ($user.DisplayName -and -not $user.DisplayName.StartsWith("Archived - ")) {
        $newDisplayName = "Archived - $($user.DisplayName)"
        Set-ADUser -Identity $user.DistinguishedName -DisplayName $newDisplayName
        Write-Host "Updated display name for: $($user.SamAccountName)" -ForegroundColor Green
    } else {
        Write-Host "Skipping (already archived or no display name): $($user.SamAccountName)" -ForegroundColor Yellow
    }
}