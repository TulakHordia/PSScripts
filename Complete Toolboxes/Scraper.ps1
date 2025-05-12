<#
.SYNOPSIS
    Inactive AD Object Management Script
.DESCRIPTION
    This script helps you find, tag, and disable inactive PCs, users, and export group memberships.
.VERSION
    AUTO_VERSION
.AUTHOR
    Benjamin Rain - Twistech
.LAST UPDATED
    AUTO_DATE
#>

# Import Active Directory module
Import-Module ActiveDirectory

# Versioning
try {
    $scriptVersion = git describe --tags --abbrev=0 2>$null
    if ([string]::IsNullOrWhiteSpace($scriptVersion)) {
        $scriptVersion = "v1.0.0"
    }
} catch {
    $scriptVersion = "v1.0.0"
}

# Global Variables
$global:thresholdDays = 30
$global:disabledDescription = "Disabled by Beni - 07/05/2025"
$domainName = (Get-ADDomain).Name
$folderPath = "C:\Twistech\Script Results"
$pcPath = "$folderPath\$domainName-InactiveComputers.csv"
$usersPath = "$folderPath\$domainName-InactiveUsers.csv"

# Create the folder if it doesn't exist
function Create-ScriptResultsFolder {
    if (-not (Test-Path -Path $folderPath)) {
        New-Item -Path $folderPath -ItemType Directory | Out-Null
}
}

# Menu function
function Show-Menu {
    Clear-Host

    # Top banner
    Write-Host "=============================================" -ForegroundColor DarkCyan
    Write-Host "      Inactive Computer and User Scraper" -ForegroundColor Green
    Write-Host "=============================================" -ForegroundColor DarkCyan
    Write-Host ""

    # Domain Info
    Write-Host ("Domain            : {0}" -f $domainName) -ForegroundColor Yellow
    Write-Host ("Current Threshold : {0} days" -f $thresholdDays) -ForegroundColor Cyan
    Write-Host ("Current Description     : {0}" -f $disabledDescription) -ForegroundColor Magenta
    Write-Host ("Script Version     : {0}" -f $scriptVersion) -ForegroundColor DarkGray

    Write-Host ""

    # Menu Options
    Write-Host "Available Actions:" -ForegroundColor White
	Write-Host "  ------ PC ------"
    Write-Host "  [1] Scrape Inactive PCs (find inactive computer accounts)"
    Write-Host "  [2] Add description to scraped PCs"
    Write-Host "  [3] Disable scraped PCs"
	Write-Host "  ------ Users ------"
    Write-Host "  [4] Scrape inactive user accounts (find inactive user accounts)"
    Write-Host "  [5] Add description to scraped users"
    Write-Host "  [6] Disable scraped users"
	Write-Host "  ------ Change Variables ------"
    Write-Host "  [7] Change current Description"
    Write-Host "  [8] Change threshold days"
    Write-Host "  ------ Groups ------"
    Write-Host "  [9] Scrape AD Groups and Members"
	Write-Host ""
    Write-Host "  [Q] Exit"

    # Bottom divider
    Write-Host ""
    Write-Host "=============================================" -ForegroundColor DarkCyan
}


# Main Program Loop
do {
    Show-Menu
    $choice = Read-Host "Enter your choice (1-9, Q to quit)"

    switch ($choice) {
        "1" { Scrape-PCs }
        "2" { Add-Description-To-PCs }
        "3" { Disable-PCs }
        "4" { Scrape-Users }
        "5" { Add-Description-To-Users }
        "6" { Disable-Users }
        "7" { Set-Description }
        "8" { Set-Threshold }
        "9" { Scrape-Groups }
        "Q" { Write-Host "Exiting..."; break }
        default { Write-Host "Invalid choice, please select a valid option." -ForegroundColor Red }
    }

    if ($choice -ne "Q") {
        Read-Host "Press Enter to continue..."
    }
}
while ($choice -ne "Q")


function Scrape-PCs {
    $global:resultsPCs = @()
    $dateThreshold = (Get-Date).AddDays(-$thresholdDays)
    
    $computers = Get-ADComputer -Filter {Enabled -eq $true} -Property LastLogonDate
    foreach ($computer in $computers) {
        $lastLogon = $computer.LastLogonDate
        if ($lastLogon -and $lastLogon -lt $dateThreshold) {
            $global:resultsPCs += $computer
        }
    }

    if ($resultsPCs.Count -gt 0) {
        Write-Host "$($resultsPCs.Count) inactive computers found."
        Create-ScriptResultsFolder
        $resultsPCs | Export-Csv -Path $pcPath -NoTypeInformation
        Write-Host "Results exported to: $pcPath"
    } else {
        Write-Host "No inactive computers found."
    }
}

function Add-Description-To-PCs {
    if (!$resultsPCs) {
        Write-Host "No PCs scraped yet. Please run option 1 first." -ForegroundColor Yellow
        return
    }

    foreach ($computer in $resultsPCs) {
        Set-ADComputer -Identity $computer.DistinguishedName -Description $disabledDescription
        Write-Host "Added description to: $($computer.Name)"
    }
    Write-Host "Finished adding descriptions."
}

function Disable-PCs {
    if (!$resultsPCs) {
        Write-Host "No PCs scraped yet. Please run option 1 first." -ForegroundColor Yellow
        return
    }

    foreach ($computer in $resultsPCs) {
        Disable-ADAccount -Identity $computer.DistinguishedName
        Write-Host "Disabled: $($computer.Name)"
    }
    Write-Host "Finished disabling computers."
}

function Scrape-Users {
    $global:resultsUsers = @()
    $dateThreshold = (Get-Date).AddDays(-$thresholdDays)

    $users = Get-ADUser -Filter {Enabled -eq $true -and LastLogonDate -lt $dateThreshold} -Property LastLogonDate
    foreach ($user in $users) {
        $global:resultsUsers += $user
    }

    if ($resultsUsers.Count -gt 0) {
        Write-Host "$($resultsUsers.Count) inactive users found."
        Create-ScriptResultsFolder
        $resultsUsers | Export-Csv -Path $usersPath -NoTypeInformation
        Write-Host "Results exported to: $usersPath"
    } else {
        Write-Host "No inactive users found."
    }
}

function Add-Description-To-Users {
    if (!$resultsUsers) {
        Write-Host "No users scraped yet. Please run option 4 first." -ForegroundColor Yellow
        return
    }

    foreach ($user in $resultsUsers) {
        Set-ADUser -Identity $user.DistinguishedName -Description $disabledDescription
        Write-Host "Added description to: $($user.Name)"
    }
    Write-Host "Finished adding descriptions."
}

function Disable-Users {
    if (!$resultsUsers) {
        Write-Host "No users scraped yet. Please run option 4 first." -ForegroundColor Yellow
        return
    }

    foreach ($user in $resultsUsers) {
        Disable-ADAccount -Identity $user.DistinguishedName
        Write-Host "Disabled: $($user.Name)"
    }
    Write-Host "Finished disabling users."
}

function Set-Description {
    $newDescription = Read-Host "Enter new description for disabled PCs and users"
    if ($newDescription) {
        $global:disabledDescription = $newDescription
        Write-Host "Description updated to: $disabledDescription"
    } else {
        Write-Host "Invalid input. Please enter a valid description." -ForegroundColor Red
    }
}
function Set-Threshold {
    $newThreshold = Read-Host "Enter the number of days for inactivity threshold"
    if ([int]::TryParse($newThreshold, [ref]$null)) {
        $global:thresholdDays = [int]$newThreshold
        Write-Host "Threshold updated to $thresholdDays days."
    } else {
        Write-Host "Invalid input. Please enter a number." -ForegroundColor Red
    }
}

function Scrape-Groups {
    $groups = Get-ADGroup -Filter * | Sort-Object Name
    $groupMembers = @()

    foreach ($group in $groups) {
        try {
            $members = Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction Stop
            foreach ($member in $members) {
                $groupMembers += [pscustomobject]@{
                    GroupName   = $group.Name
                    MemberName  = $member.Name
                    MemberType  = $member.objectClass
                    MemberDN    = $member.DistinguishedName
                }
            }
        } catch {
            # In case the group has no members or there's an access issue
            $groupMembers += [pscustomobject]@{
                GroupName   = $group.Name
                MemberName  = "(No members or access denied)"
                MemberType  = ""
                MemberDN    = ""
            }
        }
    }

    Create-ScriptResultsFolder
    $groupPath = "$folderPath\$domainName-GroupMembers.csv"
    $groupMembers | Export-Csv -Path $groupPath -NoTypeInformation
    Write-Host "Group members scraped and exported to: $groupPath" -ForegroundColor Green
}