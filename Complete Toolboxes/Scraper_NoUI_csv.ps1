<#
.SYNOPSIS
    Inactive AD Object Management Script
.DESCRIPTION
    This script helps you find, tag, disable, and relocate inactive PCs and users,
    and create the OUs that disabled objects are moved into.
.VERSION
    v1.1.0
.AUTHOR
    Benjamin Rain
.LAST UPDATED
    2026-06-29
#>

Import-Module ActiveDirectory

# ── Global Variables ───────────────────────────────────────────────────────────
$global:thresholdDays            = 30
$global:disabledDescription      = "Disabled by Beni - 29/06/2026"
$global:disabledUsersOUName      = "Disabled Users"
$global:disabledComputersOUName  = "Disabled Computers"

$domainName                      = (Get-ADDomain).Name
$domainDN                        = (Get-ADDomain).DistinguishedName
$folderPath                      = "C:\Informat\Script Results"
$scriptVersion                   = "v1.1.0"
$pcPath                          = "$folderPath\$domainName-InactiveComputers.csv"
$usersPath                       = "$folderPath\$domainName-InactiveUsers.csv"

Write-Host "Current Inactivity Threshold : $thresholdDays days" -ForegroundColor Cyan
Write-Host "Current Disabled Description : '$disabledDescription'" -ForegroundColor Cyan

# ── Helpers ────────────────────────────────────────────────────────────────────
function New-ScriptResultsFolder {
    if (-not (Test-Path -Path $folderPath)) {
        New-Item -Path $folderPath -ItemType Directory | Out-Null
    }
}

function Pause-Menu {
    Write-Host ""
    Write-Host "Press any key to return to menu..." -ForegroundColor DarkGray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Ensures an OU exists directly under $ParentDN, creating it if needed.
# Returns the OU's DistinguishedName, or $null on failure.
function Ensure-OU {
    param(
        [Parameter(Mandatory)] [string]$OUName,
        [Parameter(Mandatory)] [string]$ParentDN
    )

    $ouDN = "OU=$OUName,$ParentDN"

    $existing = Get-ADOrganizationalUnit -Filter "Name -eq '$OUName'" `
                    -SearchBase $ParentDN -SearchScope OneLevel -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "OU already exists: $($existing.DistinguishedName)" -ForegroundColor DarkGray
        return $existing.DistinguishedName
    }

    try {
        New-ADOrganizationalUnit -Name $OUName -Path $ParentDN `
            -ProtectedFromAccidentalDeletion $true -ErrorAction Stop
        Write-Host "Created OU: $ouDN" -ForegroundColor Green
        return $ouDN
    } catch {
        Write-Host "Failed to create OU '$OUName': $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

# ── OU Functions ───────────────────────────────────────────────────────────────
function New-DisabledOUs {
    Write-Host "Ensuring disabled-object OUs exist under: $domainDN" -ForegroundColor Cyan
    Ensure-OU -OUName $global:disabledComputersOUName -ParentDN $domainDN | Out-Null
    Ensure-OU -OUName $global:disabledUsersOUName     -ParentDN $domainDN | Out-Null
}

# ── PC Functions ───────────────────────────────────────────────────────────────
function GetPCs {
    $global:resultsPCs = @()
    $dateThreshold = (Get-Date).AddDays(-$global:thresholdDays)
    $computers = Get-ADComputer -Filter { Enabled -eq $true } -Properties LastLogonDate

    foreach ($computer in $computers) {
        if ($computer.LastLogonDate -and $computer.LastLogonDate -lt $dateThreshold) {
            $global:resultsPCs += $computer
        }
    }

    if ($global:resultsPCs.Count -gt 0) {
        Write-Host "$($global:resultsPCs.Count) inactive computer(s) found." -ForegroundColor White
        New-ScriptResultsFolder
        $global:resultsPCs | Select-Object Name, DistinguishedName, LastLogonDate |
            Export-Csv -Path $pcPath -NoTypeInformation
        Write-Host "Saved to: $pcPath" -ForegroundColor Green
    } else {
        Write-Host "No inactive PCs found." -ForegroundColor Yellow
    }
}

function Add-Description-To-PCs {
    if (-not $global:resultsPCs) {
        Write-Host "No PCs scraped yet. Run option 1 first." -ForegroundColor Red
        return
    }
    foreach ($computer in $global:resultsPCs) {
        Set-ADComputer -Identity $computer.DistinguishedName -Description $global:disabledDescription
        Write-Host "Description added: $($computer.Name)" -ForegroundColor White
    }
}

function Disable-PCs {
    if (-not $global:resultsPCs) {
        Write-Host "No PCs scraped yet. Run option 1 first." -ForegroundColor Red
        return
    }
    foreach ($computer in $global:resultsPCs) {
        Disable-ADAccount -Identity $computer.DistinguishedName
        Write-Host "Disabled: $($computer.Name)" -ForegroundColor White
    }
}

function Move-PCs-ToDisabledOU {
    if (-not $global:resultsPCs) {
        Write-Host "No PCs scraped yet. Run option 1 first." -ForegroundColor Red
        return
    }

    $targetOU = Ensure-OU -OUName $global:disabledComputersOUName -ParentDN $domainDN
    if (-not $targetOU) {
        Write-Host "Target OU unavailable. Aborting move." -ForegroundColor Red
        return
    }

    foreach ($computer in $global:resultsPCs) {
        try {
            Move-ADObject -Identity $computer.DistinguishedName -TargetPath $targetOU -ErrorAction Stop
            Write-Host "Moved: $($computer.Name)" -ForegroundColor White
        } catch {
            Write-Host "Failed to move $($computer.Name): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# ── User Functions ─────────────────────────────────────────────────────────────
function GetUsers {
    $global:resultsUsers = @()
    $dateThreshold = (Get-Date).AddDays(-$global:thresholdDays)
    $users = Get-ADUser -Filter { Enabled -eq $true } -Properties LastLogonDate |
             Where-Object { $_.LastLogonDate -and $_.LastLogonDate -lt $dateThreshold }

    foreach ($user in $users) {
        $global:resultsUsers += $user
    }

    if ($global:resultsUsers.Count -gt 0) {
        Write-Host "$($global:resultsUsers.Count) inactive user(s) found." -ForegroundColor White
        New-ScriptResultsFolder
        $global:resultsUsers | Select-Object Name, SamAccountName, DistinguishedName, LastLogonDate |
            Export-Csv -Path $usersPath -NoTypeInformation
        Write-Host "Saved to: $usersPath" -ForegroundColor Green
    } else {
        Write-Host "No inactive users found." -ForegroundColor Yellow
    }
}

function Add-Description-To-Users {
    if (-not $global:resultsUsers) {
        Write-Host "No users scraped yet. Run option 5 first." -ForegroundColor Red
        return
    }
    foreach ($user in $global:resultsUsers) {
        Set-ADUser -Identity $user.DistinguishedName -Description $global:disabledDescription
        Write-Host "Description added: $($user.Name)" -ForegroundColor White
    }
}

function Disable-Users {
    if (-not $global:resultsUsers) {
        Write-Host "No users scraped yet. Run option 5 first." -ForegroundColor Red
        return
    }
    foreach ($user in $global:resultsUsers) {
        Disable-ADAccount -Identity $user.DistinguishedName
        Write-Host "Disabled: $($user.Name)" -ForegroundColor White
    }
}

function Move-Users-ToDisabledOU {
    if (-not $global:resultsUsers) {
        Write-Host "No users scraped yet. Run option 5 first." -ForegroundColor Red
        return
    }

    $targetOU = Ensure-OU -OUName $global:disabledUsersOUName -ParentDN $domainDN
    if (-not $targetOU) {
        Write-Host "Target OU unavailable. Aborting move." -ForegroundColor Red
        return
    }

    foreach ($user in $global:resultsUsers) {
        try {
            Move-ADObject -Identity $user.DistinguishedName -TargetPath $targetOU -ErrorAction Stop
            Write-Host "Moved: $($user.Name)" -ForegroundColor White
        } catch {
            Write-Host "Failed to move $($user.Name): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# ── Settings Functions ─────────────────────────────────────────────────────────
function Set-Description {
    $newDescription = Read-Host "Enter new description for disabled objects"
    if (-not [string]::IsNullOrWhiteSpace($newDescription)) {
        $global:disabledDescription = $newDescription
        Write-Host "Description updated to: $global:disabledDescription" -ForegroundColor Cyan
    } else {
        Write-Host "Invalid input. Description unchanged." -ForegroundColor Red
    }
}

function Set-Threshold {
    $newThreshold = Read-Host "Enter the number of days for inactivity threshold"
    $parsed = 0
    if ([int]::TryParse($newThreshold, [ref]$parsed) -and $parsed -gt 0) {
        $global:thresholdDays = $parsed
        Write-Host "Threshold updated to: $global:thresholdDays days" -ForegroundColor Cyan
    } else {
        Write-Host "Invalid input. Please enter a positive number." -ForegroundColor Red
    }
}

# ── Menu ───────────────────────────────────────────────────────────────────────
function Show-Menu {
    Clear-Host
    Write-Host "=================================" -ForegroundColor Cyan
    Write-Host "  Inactive AD Object Manager $scriptVersion" -ForegroundColor Yellow
    Write-Host "=================================" -ForegroundColor Cyan
    Write-Host "  Threshold   : $global:thresholdDays days" -ForegroundColor DarkCyan
    Write-Host "  Description : $global:disabledDescription" -ForegroundColor DarkCyan
    Write-Host "  Disabled PCs OU   : $global:disabledComputersOUName" -ForegroundColor DarkCyan
    Write-Host "  Disabled Users OU : $global:disabledUsersOUName" -ForegroundColor DarkCyan
    Write-Host "---------------------------------" -ForegroundColor DarkGray
    Write-Host "  -- Computers --"
    Write-Host "  1. Scrape Inactive PCs"
    Write-Host "  2. Add Description to PCs"
    Write-Host "  3. Disable Scraped PCs"
    Write-Host "  4. Move Scraped PCs -> Disabled Computers OU"
    Write-Host "  -- Users --"
    Write-Host "  5. Scrape Inactive Users"
    Write-Host "  6. Add Description to Users"
    Write-Host "  7. Disable Scraped Users"
    Write-Host "  8. Move Scraped Users -> Disabled Users OU"
    Write-Host "  -- Setup & Settings --"
    Write-Host "  9.  Create Disabled OUs"
    Write-Host "  10. Set Disabled Description"
    Write-Host "  11. Set Inactivity Threshold"
    Write-Host "  0.  Exit"
    Write-Host "=================================" -ForegroundColor Cyan
}

# ── Main Loop ──────────────────────────────────────────────────────────────────
while ($true) {
    Show-Menu
    $choice = Read-Host "Select an option"
    switch ($choice) {
        '1'  { GetPCs }
        '2'  { Add-Description-To-PCs }
        '3'  { Disable-PCs }
        '4'  { Move-PCs-ToDisabledOU }
        '5'  { GetUsers }
        '6'  { Add-Description-To-Users }
        '7'  { Disable-Users }
        '8'  { Move-Users-ToDisabledOU }
        '9'  { New-DisabledOUs }
        '10' { Set-Description }
        '11' { Set-Threshold }
        '0'  { exit }
        default { Write-Host "Invalid selection." -ForegroundColor Red }
    }
    Pause-Menu
}