<#
.SYNOPSIS
    Inactive AD Object Management Script
.DESCRIPTION
    This script helps you find, tag, and disable inactive PCs, users, and export group memberships.
.VERSION
    v1.0.2
.AUTHOR
    Benjamin Rain - Twistech
.LAST UPDATED
    2025-07-03
#>

# Versioning
$scriptVersion = "v1.0.2"

# Load necessary assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Import AD Module
Import-Module ActiveDirectory

# Global Variables
$global:thresholdDays = 30
$global:disabledDescription = "Disabled by Beni - 07/05/2025"
$domainName = (Get-ADDomain).Name
$folderPath = "C:\Twistech\Script Results"
$pcPath = "$folderPath\$domainName-InactiveComputers.csv"
$usersPath = "$folderPath\$domainName-InactiveUsers.csv"
$groupPath = "$folderPath\$domainName-GroupMembers.csv"

function New-ScriptResultsFolder {
    if (-not (Test-Path -Path $folderPath)) {
        New-Item -Path $folderPath -ItemType Directory | Out-Null
    }
}

# GUI
$form = New-Object System.Windows.Forms.Form
$form.Text = "Inactive AD Object Management"
$form.Size = New-Object System.Drawing.Size(450, 500)
$form.StartPosition = "CenterScreen"

$title = New-Object System.Windows.Forms.Label
$title.Text = "Inactive AD Object Manager ($scriptVersion)"
$title.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$title.Size = New-Object System.Drawing.Size(400, 30)
$title.Location = New-Object System.Drawing.Point(20, 20)
$form.Controls.Add($title)

$buttons = @(
    @{Text = "1. Scrape Inactive PCs"; Action = { Get-InactivePCs }},
    @{Text = "2. Add Description to PCs"; Action = { Add-Description-To-PCs }},
    @{Text = "3. Disable Scraped PCs"; Action = { Disable-PCs }},
    @{Text = "4. Scrape Inactive Users"; Action = { Get-InactiveUsers }},
    @{Text = "5. Add Description to Users"; Action = { Add-Description-To-Users }},
    @{Text = "6. Disable Scraped Users"; Action = { Disable-Users }},
    @{Text = "7. Set Disabled Description"; Action = { Set-Description }},
    @{Text = "8. Set Inactivity Threshold"; Action = { Set-Threshold }},
    @{Text = "9. Scrape Groups"; Action = { Get-GroupMembers }}
)

for ($i = 0; $i -lt $buttons.Count; $i++) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $buttons[$i].Text
    $btn.Size = New-Object System.Drawing.Size(380, 30)
    $btn.Location = New-Object System.Drawing.Point(30, (70 + ($i * 40)))
    $btn.Add_Click($buttons[$i].Action)
    $form.Controls.Add($btn)
}

$form.Topmost = $true
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()

function Get-InactivePCs {
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
        $resultsPCs | Select-Object Name, DistinguishedName, LastLogonDate | Export-Csv -Path $pcPath -NoTypeInformation
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

function Get-InactiveUsers {
    $global:resultsUsers = @()
    $dateThreshold = (Get-Date).AddDays(-$thresholdDays)

    $users = Get-ADUser -Filter {Enabled -eq $true -and LastLogonDate -lt $dateThreshold} -Property LastLogonDate
    foreach ($user in $users) {
        $global:resultsUsers += $user
    }

    if ($resultsUsers.Count -gt 0) {
        Write-Host "$($resultsUsers.Count) inactive users found."
        New-ScriptResultsFolder
        $resultsUsers | Select-Object Name, SamAccountName, DistinguishedName, LastLogonDate | Export-Csv -Path $usersPath -NoTypeInformation
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

function Get-GroupMembers {
    $groups = Get-ADGroup -Filter * | Sort-Object Name
    $groupMembers = @()

    foreach ($group in $groups) {
        try {
            $members = Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction Stop
            foreach ($member in $members) {
                if ($member.objectClass -eq 'user') {
                    $groupMembers += [pscustomobject]@{
                        GroupName   = $group.Name
                        MemberName  = $member.Name
                        MemberType  = $member.objectClass
                        MemberDN    = $member.DistinguishedName
                    }
                }
            }
        } catch {
            $groupMembers += [pscustomobject]@{
                GroupName   = $group.Name
                MemberName  = "(No members or access denied)"
                MemberType  = ""
                MemberDN    = ""
            }
        }
    }

    New-ScriptResultsFolder
    $groupMembers | Export-Csv -Path $groupPath -NoTypeInformation
    Write-Host "Group members (Users only) scraped and exported to: $groupPath" -ForegroundColor Green
}
