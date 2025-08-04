# Import Active Directory module
Import-Module ActiveDirectory

# Path to exclusion list
$exclusionPath = "C:\Twistech\Scripts\excluded_users.csv"
$excludedUsers = @()

# Load exclusions from CSV
if (Test-Path $exclusionPath) {
    $excludedUsers = Import-Csv -Path $exclusionPath | Select-Object -ExpandProperty SamAccountName
} else {
    Write-Host "Exclusion file not found at: $exclusionPath" -ForegroundColor Red
    exit
}

# Get all disabled users
$disabledUsers = Get-ADUser -Filter {Enabled -eq $false} -Properties DisplayName, Name, SamAccountName, EmailAddress, msExchHideFromAddressLists

foreach ($user in $disabledUsers) {
    $username = $user.SamAccountName

    # Always apply HideFromAddressLists to ALL disabled users
    try {
        Set-ADUser -Identity $user.DistinguishedName -Replace @{msExchHideFromAddressLists = $true}
        Write-Host "Set msExchHideFromAddressLists=TRUE for: $username" -ForegroundColor Blue
    } catch {
        Write-Host "Failed to set HideFromAddressLists for: $username" -ForegroundColor Red
    }

    # Skip if excluded
    if ($excludedUsers -contains $username) {
        Write-Host "Excluded by CSV: $username" -ForegroundColor Cyan
        continue
    }

    # Skip if no email
    if ([string]::IsNullOrWhiteSpace($user.EmailAddress)) {
        Write-Host "Skipped (no email): $username" -ForegroundColor Gray
        continue
    }

    # Update DisplayName
    if ($user.DisplayName -and -not $user.DisplayName.StartsWith("Archived - ")) {
        $newDisplayName = "Archived - $($user.DisplayName)"
        Set-ADUser -Identity $user.DistinguishedName -DisplayName $newDisplayName
        Write-Host "Updated DisplayName: $username" -ForegroundColor Green
    } else {
        Write-Host "DisplayName already archived or missing: $username" -ForegroundColor DarkYellow
    }

    # Update Name (Full Name)
    if ($user.Name -and -not $user.Name.StartsWith("Archived - ")) {
        $newName = "Archived - $($user.Name)"
        Rename-ADObject -Identity $user.DistinguishedName -NewName $newName
        Write-Host "Updated Full Name: $username" -ForegroundColor Green
    } else {
        Write-Host "Full Name already archived or missing: $username" -ForegroundColor DarkYellow
    }
}
