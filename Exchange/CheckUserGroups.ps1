# Import modules
Import-Module Microsoft.Graph.Users
Import-Module ExchangeOnlineManagement

# Connect to Microsoft Graph
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.Read.All", "GroupMember.Read.All"

# Connect to Exchange Online
Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnline -ShowProgress:$false

# Prompt for user
$userUPN = Read-Host "Enter the user's email (UPN)"

# Get user object
$user = Get-MgUser -UserId $userUPN -ErrorAction SilentlyContinue

if (-not $user) {
    Write-Host "User not found in Microsoft 365." -ForegroundColor Red
    return
}

# Check block status
$isBlocked = if ($user.AccountEnabled -eq $false) { "Yes" } else { "No" }

# Get group membership (Exchange Distribution Groups only)
$groups = Get-DistributionGroup -ResultSize Unlimited | Where-Object {
    (Get-DistributionGroupMember -Identity $_.Identity -ResultSize Unlimited -ErrorAction SilentlyContinue |
        Where-Object { $_.PrimarySmtpAddress -eq $userUPN })
} | Select-Object -ExpandProperty DisplayName

# Output
Write-Host "`nResults for: $userUPN" -ForegroundColor Yellow
Write-Host "Blocked from sign-in: $isBlocked" -ForegroundColor Green

Write-Host "Distribution Groups:"
if ($groups) {
    $groups | ForEach-Object { Write-Host " - $_" }
} else {
    Write-Host " - None"
}