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
$scriptVersion = "v1.0.1"

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

# New top-level menu function with nested submenus
# Load necessary assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Import AD Module
Import-Module ActiveDirectory

# Version and Global Variables
$scriptVersion = "v1.0.1"
$global:thresholdDays = 30
$global:disabledDescription = "Disabled by Beni - 07/05/2025"
$domainName = (Get-ADDomain).Name
$folderPath = "C:\Twistech\Script Results"
$pcPath = "$folderPath\$domainName-InactiveComputers.csv"
$usersPath = "$folderPath\$domainName-InactiveUsers.csv"

function Create-ScriptResultsFolder {
    if (-not (Test-Path -Path $folderPath)) {
        New-Item -Path $folderPath -ItemType Directory | Out-Null
    }
}

# AD Functions (from your existing logic)
function Scrape-PCs { ... }
function Add-Description-To-PCs { ... }
function Disable-PCs { ... }
function Scrape-Users { ... }
function Add-Description-To-Users { ... }
function Disable-Users { ... }
function Set-Description { ... }
function Set-Threshold { ... }
function Scrape-Groups { ... }

# Create Main Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Inactive AD Object Management"
$form.Size = New-Object System.Drawing.Size(450, 500)
$form.StartPosition = "CenterScreen"

# Title Label
$title = New-Object System.Windows.Forms.Label
$title.Text = "Inactive AD Object Manager ($scriptVersion)"
$title.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$title.Size = New-Object System.Drawing.Size(400, 30)
$title.Location = New-Object System.Drawing.Point(20, 20)
$form.Controls.Add($title)

# PC Actions
$btnPC1 = New-Object System.Windows.Forms.Button
$btnPC1.Text = "1. Scrape Inactive PCs"
$btnPC1.Size = New-Object System.Drawing.Size(380, 30)
$btnPC1.Location = New-Object System.Drawing.Point(30, 70)
$btnPC1.Add_Click({ Scrape-PCs })
$form.Controls.Add($btnPC1)

$btnPC2 = New-Object System.Windows.Forms.Button
$btnPC2.Text = "2. Add Description to PCs"
$btnPC2.Size = New-Object System.Drawing.Size(380, 30)
$btnPC2.Location = New-Object System.Drawing.Point(30, 110)
$btnPC2.Add_Click({ Add-Description-To-PCs })
$form.Controls.Add($btnPC2)

$btnPC3 = New-Object System.Windows.Forms.Button
$btnPC3.Text = "3. Disable Scraped PCs"
$btnPC3.Size = New-Object System.Drawing.Size(380, 30)
$btnPC3.Location = New-Object System.Drawing.Point(30, 150)
$btnPC3.Add_Click({ Disable-PCs })
$form.Controls.Add($btnPC3)

# User Actions
$btnUser1 = New-Object System.Windows.Forms.Button
$btnUser1.Text = "4. Scrape Inactive Users"
$btnUser1.Size = New-Object System.Drawing.Size(380, 30)
$btnUser1.Location = New-Object System.Drawing.Point(30, 200)
$btnUser1.Add_Click({ Scrape-Users })
$form.Controls.Add($btnUser1)

$btnUser2 = New-Object System.Windows.Forms.Button
$btnUser2.Text = "5. Add Description to Users"
$btnUser2.Size = New-Object System.Drawing.Size(380, 30)
$btnUser2.Location = New-Object System.Drawing.Point(30, 240)
$btnUser2.Add_Click({ Add-Description-To-Users })
$form.Controls.Add($btnUser2)

$btnUser3 = New-Object System.Windows.Forms.Button
$btnUser3.Text = "6. Disable Scraped Users"
$btnUser3.Size = New-Object System.Drawing.Size(380, 30)
$btnUser3.Location = New-Object System.Drawing.Point(30, 280)
$btnUser3.Add_Click({ Disable-Users })
$form.Controls.Add($btnUser3)

# Variable Settings
$btnVar1 = New-Object System.Windows.Forms.Button
$btnVar1.Text = "7. Set Disabled Description"
$btnVar1.Size = New-Object System.Drawing.Size(380, 30)
$btnVar1.Location = New-Object System.Drawing.Point(30, 330)
$btnVar1.Add_Click({ Set-Description })
$form.Controls.Add($btnVar1)

$btnVar2 = New-Object System.Windows.Forms.Button
$btnVar2.Text = "8. Set Inactivity Threshold"
$btnVar2.Size = New-Object System.Drawing.Size(380, 30)
$btnVar2.Location = New-Object System.Drawing.Point(30, 370)
$btnVar2.Add_Click({ Set-Threshold })
$form.Controls.Add($btnVar2)

# Group Actions
$btnGroup = New-Object System.Windows.Forms.Button
$btnGroup.Text = "9. Scrape Groups"
$btnGroup.Size = New-Object System.Drawing.Size(380, 30)
$btnGroup.Location = New-Object System.Drawing.Point(30, 410)
$btnGroup.Add_Click({ Scrape-Groups })
$form.Controls.Add($btnGroup)

# Show the Form
$form.Topmost = $true
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()



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