<#
.SYNOPSIS
    Inactive AD Object Management Script
.DESCRIPTION
    This script helps you find, tag, and disable inactive PCs, users, and export group memberships.
.VERSION
    v1.0.1
.AUTHOR
    Benjamin Rain - Twistech
.LAST UPDATED
    2025-05-15
#>

# Load required modules
Import-Module ActiveDirectory

# Check and import ImportExcel
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Installing ImportExcel module..." -ForegroundColor Cyan
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}
Import-Module ImportExcel

# Global Variables
$global:thresholdDays = 30
$global:disabledDescription = "Disabled by Beni - 07/05/2025"
$domainName = (Get-ADDomain).Name
$folderPath = "C:\Twistech\Script Results"
$scriptVersion = "v1.0.1"
$pcPath = "$folderPath\$domainName-InactiveComputers.xlsx"
$usersPath = "$folderPath\$domainName-InactiveUsers.xlsx"
$groupPath = "$folderPath\$domainName-GroupMembers.xlsx"

function New-ScriptResultsFolder {
    if (-not (Test-Path -Path $folderPath)) {
        New-Item -Path $folderPath -ItemType Directory | Out-Null
    }
}

function GetPCs {
    $global:resultsPCs = @()
    $dateThreshold = (Get-Date).AddDays(-$thresholdDays)
    $computers = Get-ADComputer -Filter {Enabled -eq $true} -Property LastLogonDate
    foreach ($computer in $computers) {
        if ($computer.LastLogonDate -and $computer.LastLogonDate -lt $dateThreshold) {
            $global:resultsPCs += $computer
        }
    }

    if ($resultsPCs.Count -gt 0) {
        Write-Host "$($resultsPCs.Count) inactive computers found."
        New-ScriptResultsFolder
        $resultsPCs | Export-Excel -Path $pcPath -WorksheetName "Inactive PCs" -AutoSize
        Write-Host "Scraped inactive PCs and saved to: $pcPath" -ForegroundColor Green
    } else {
        Write-Host "No inactive PCs found." -ForegroundColor Yellow
    }
}

function Add-Description-To-PCs {
    if (!$resultsPCs) {
        Write-Host "No PCs scraped yet. Please run option 1 first." -ForegroundColor Red
        return
    }

    foreach ($computer in $resultsPCs) {
        Set-ADComputer -Identity $computer.DistinguishedName -Description $disabledDescription
        Write-Host "Added description to: $($computer.Name)"
    }
}

function Disable-PCs {
    if (!$resultsPCs) {
        Write-Host "No PCs scraped yet. Please run option 1 first." -ForegroundColor Red
        return
    }

    foreach ($computer in $resultsPCs) {
        Disable-ADAccount -Identity $computer.DistinguishedName
        Write-Host "Disabled: $($computer.Name)"
    }
}

function GetUsers {
    $global:resultsUsers = @()
    $dateThreshold = (Get-Date).AddDays(-$thresholdDays)
    $users = Get-ADUser -Filter {Enabled -eq $true -and LastLogonDate -lt $dateThreshold} -Property LastLogonDate
    foreach ($user in $users) {
        $global:resultsUsers += $user
    }

    if ($resultsUsers.Count -gt 0) {
        Write-Host "$($resultsUsers.Count) inactive users found."
        New-ScriptResultsFolder
        $resultsUsers | Export-Excel -Path $usersPath -WorksheetName "Inactive Users" -AutoSize
        Write-Host "Scraped inactive Users and saved to: $usersPath" -ForegroundColor Green
    } else {
        Write-Host "No inactive users found." -ForegroundColor Yellow
    }
}

function Add-Description-To-Users {
    if (!$resultsUsers) {
        Write-Host "No Users scraped yet. Please run option 4 first." -ForegroundColor Red
        return
    }

    foreach ($user in $resultsUsers) {
        Set-ADUser -Identity $user.DistinguishedName -Description $disabledDescription
        Write-Host "Added description to: $($user.Name)"
    }
}

function Disable-Users {
    if (!$resultsUsers) {
        Write-Host "No Users scraped yet. Please run option 4 first." -ForegroundColor Red
        return
    }

    foreach ($user in $resultsUsers) {
        Disable-ADAccount -Identity $user.DistinguishedName
        Write-Host "Disabled: $($user.Name)"
    }
}

function Set-Description {
    $newDescription = Read-Host "Enter new description for disabled PCs and users"
    if ($newDescription) {
        $global:disabledDescription = $newDescription
        Write-Host "Description updated to: $disabledDescription" -ForegroundColor Cyan
    } else {
        Write-Host "Invalid input. Please enter a valid description" -ForegroundColor Red
    }
}

function Set-Threshold {
    $newThreshold = Read-Host "Enter the number of days for inactivity threshold"
    if ([int]::TryParse($newThreshold, [ref]$null)) {
        $global:thresholdDays = [int]$newThreshold
        Write-Host "Threshold updated to: $thresholdDays" -ForegroundColor Cyan
    } else {
        Write-Host "Invalid input. Please enter a number" -ForegroundColor Red
    }
}

function GetGroups {
    $groups = Get-ADGroup -Filter * -Properties * | Sort-Object Name
    $groupMembers = @()

    foreach ($group in $groups) {
        try {
            $members = Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction Stop
            foreach ($member in $members) {
                if ($member.objectClass -eq 'user') {
                    $groupMembers += [pscustomobject]@{
                        GroupName        = $group.Name
                        MemberName       = $member.Name
                        MemberType       = $member.objectClass
                        MemberUsername   = $member.SamAccountName
                        GroupDescription = $group.description
                    }
                }
            }
        } catch {
            $groupMembers += [pscustomobject]@{
                GroupName        = $group.Name
                MemberUsername   = ""
                MemberName       = "(No members or access denied)"
                MemberType       = ""
                GroupDescription = ""
            }
        }
    }

    New-ScriptResultsFolder
    $groupMembers | Export-Excel -Path $groupPath -WorksheetName "Group Members" -AutoSize
    Write-Host "Group members (Users only) exported to: $groupPath" -ForegroundColor Green
}

# Menu
function Show-Menu {
    Clear-Host
    Write-Host "==============================" -ForegroundColor Cyan
    Write-Host " Inactive AD Object Manager v$scriptVersion " -ForegroundColor Yellow
    Write-Host "==============================" -ForegroundColor Cyan
    Write-Host "1. Scrape Inactive PCs"
    Write-Host "2. Add Description to PCs"
    Write-Host "3. Disable Scraped PCs"
    Write-Host "4. Scrape Inactive Users"
    Write-Host "5. Add Description to Users"
    Write-Host "6. Disable Scraped Users"
    Write-Host "7. Set Disabled Description"
    Write-Host "8. Set Inactivity Threshold"
    Write-Host "9. Scrape Groups"
    Write-Host "0. Exit"
    Write-Host "=============================="
}

# Main Loop
while ($true) {
    Show-Menu
    $choice = Read-Host "Select an option"
    switch ($choice) {
        '1' { GetPCs }
        '2' { Add-Description-To-PCs }
        '3' { Disable-PCs }
        '4' { GetUsers }
        '5' { Add-Description-To-Users }
        '6' { Disable-Users }
        '7' { Set-Description }
        '8' { Set-Threshold }
        '9' { GetGroups }
        '0' { break }
        default { Write-Host "Invalid selection." -ForegroundColor Red }
    }
    Pause
}
