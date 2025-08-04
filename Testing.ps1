# Load required module
try {
    Import-Module ExchangeOnlineManagement -Force -ErrorAction Stop
} catch {
    Write-Host "ExchangeOnlineManagement module not found. Attempting to install..."
    Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser
    Import-Module ExchangeOnlineManagement -Force
}

# Connect to Exchange Online
$userUPN = Read-Host "Enter your Microsoft 365 email (UPN)"
if (-not $userUPN) {
    Write-Host "No email entered. Exiting..."
    exit
}

try {
    Connect-ExchangeOnline -UserPrincipalName $userUPN -ShowBanner:$false -ErrorAction Stop
    Write-Host "`nConnected to Exchange Online as $userUPN`n" -ForegroundColor Green
} catch {
    Write-Error "Error connecting to Exchange Online: $_"
    exit
}

# Prompt for calendar permission inputs
$owner = Read-Host "Enter the calendar owner's email address"
$user = Read-Host "Enter the email address of the user to grant access"
$permOptions = @("AvailabilityOnly", "LimitedDetails", "Reviewer", "Editor", "Owner")

do {
    $perm = Read-Host "Enter the permission level (`"AvailabilityOnly`", `"LimitedDetails`", `"Reviewer`", `"Editor`", `"Owner`")"
} while ($perm -notin $permOptions)

Write-Host "`nLooking for calendar folders for $owner..."

try {
    $calendarFoldersEN = Get-MailboxFolderStatistics -Identity $owner | Where-Object { $_.FolderType -eq "Calendar" }
    $calendarFoldersHEB = Get-MailboxFolderStatistics -Identity $owner | Where-Object { $_.FolderType -eq "לוח שנה" }

    $calendarFolders = $calendarFoldersEN + $calendarFoldersHEB

    if ($calendarFolders.Count -eq 0) {
        Write-Host "No calendar folders found for $owner" -ForegroundColor Red
        exit
    }

    $calendarFound = $false

    foreach ($folder in $calendarFolders) {
        $folderPath = $folder.FolderPath -replace "\\", "/"
        Write-Host "Attempting to assign '$perm' to '$user' for folder: $folderPath"

        try {
            Add-MailboxFolderPermission -Identity "$owner`:$folderPath" -User $user -AccessRights $perm -ErrorAction Stop
            Write-Host "Successfully added $perm permission to $user on $folderPath" -ForegroundColor Green
            $calendarFound = $true
            break
        } catch {
            Write-Warning "Failed to assign permission on $folderPath: $_"
        }
    }

    if (-not $calendarFound) {
        Write-Host "Failed to assign calendar permission. No matching folder found or all attempts failed." -ForegroundColor Yellow
    } else {
        $logPath = "$env:USERPROFILE\calendar_permission_log.txt"
        "Granted $perm permission to $user for $owner calendar ($folderPath) - $(Get-Date)" | Out-File -FilePath $logPath -Encoding UTF8 -Append
        Write-Host "`nLog saved to: $logPath"
    }

} catch {
    Write-Error "An error occurred while processing calendar permissions: $_"
}
