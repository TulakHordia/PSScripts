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

# Inactive PCs
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
        [System.Windows.Forms.MessageBox]::Show("Scraped inactive PCs and saved to:`n$pcPath", "Export Complete", [System.Windows.Forms.MessageBoxButtons]::OK)
    } else {
        [System.Windows.Forms.MessageBox]::Show("No inactive PCs found", "Nothing Exported", [System.Windows.Forms.MessageBoxButtons]::OK)
    }
}

function Add-Description-To-PCs {
    if (!$resultsPCs) {
        [System.Windows.Forms.MessageBox]::Show("No PCs scraped yet. Please run option 1 first.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    foreach ($computer in $resultsPCs) {
        Set-ADComputer -Identity $computer.DistinguishedName -Description $disabledDescription
        Write-Host "Added description to: $($computer.Name)"
    }
    [System.Windows.Forms.MessageBox]::Show("Added Descriptions", "Description", [System.Windows.Forms.MessageBoxButtons]::OK)
}

function Disable-PCs {
    if (!$resultsPCs) {
        [System.Windows.Forms.MessageBox]::Show("No PCs scraped yet. Please run option 1 first.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    foreach ($computer in $resultsPCs) {
        Disable-ADAccount -Identity $computer.DistinguishedName
        Write-Host "Disabled: $($computer.Name)"
    }
    [System.Windows.Forms.MessageBox]::Show("Disabled PCs", "Disable", [System.Windows.Forms.MessageBoxButtons]::OK)
}

# Inactive Users
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
        [System.Windows.Forms.MessageBox]::Show("Scraped inactive Users and saved to:`n$usersPath", "Export Complete", [System.Windows.Forms.MessageBoxButtons]::OK)
    } else {
        [System.Windows.Forms.MessageBox]::Show("No inactive Users found", "Nothing Exported", [System.Windows.Forms.MessageBoxButtons]::OK)
    }
}

function Add-Description-To-Users {
    if (!$resultsUsers) {
        [System.Windows.Forms.MessageBox]::Show("No Users scraped yet. Please run option 4 first.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    foreach ($user in $resultsUsers) {
        Set-ADUser -Identity $user.DistinguishedName -Description $disabledDescription
        Write-Host "Added description to: $($user.Name)"
    }
    [System.Windows.Forms.MessageBox]::Show("Added Descriptions", "Description", [System.Windows.Forms.MessageBoxButtons]::OK)
}

function Disable-Users {
    if (!$resultsUsers) {
        [System.Windows.Forms.MessageBox]::Show("No Users scraped yet. Please run option 4 first.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    foreach ($user in $resultsUsers) {
        Disable-ADAccount -Identity $user.DistinguishedName
        Write-Host "Disabled: $($user.Name)"
    }
    [System.Windows.Forms.MessageBox]::Show("Disabled Users", "Disable", [System.Windows.Forms.MessageBoxButtons]::OK)
}

function Set-Description {
    $newDescription = Read-Host "Enter new description for disabled PCs and users"
    if ($newDescription) {
        $global:disabledDescription = $newDescription
        [System.Windows.Forms.MessageBox]::Show("Description updated to: `n$disabledDescription", "Description Updated", [System.Windows.Forms.MessageBoxButtons]::OK)
    } else {
        [System.Windows.Forms.MessageBox]::Show("Invalid input. Please enter a valid description", "Error", [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function Set-Threshold {
    $newThreshold = Read-Host "Enter the number of days for inactivity threshold"
    if ([int]::TryParse($newThreshold, [ref]$null)) {
        $global:thresholdDays = [int]$newThreshold
        [System.Windows.Forms.MessageBox]::Show("Threshold updated to: `n$thresholdDays", "Threshold Updated", [System.Windows.Forms.MessageBoxButtons]::OK)
    } else {
        [System.Windows.Forms.MessageBox]::Show("Invalid input. Please enter a number", "Error", [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error)
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
                        GroupName   = $group.Name
                        MemberName  = $member.Name
                        MemberType  = $member.objectClass
                        MemberUsername = $member.SamAccountName
                        GroupDescription = $group.description
                    }
                }
            }
        } catch {
            $groupMembers += [pscustomobject]@{
                GroupName   = $group.Name
                MemberUsername = ""
                MemberName  = "(No members or access denied)"
                MemberType  = ""
                GroupDescription = ""
            }
        }
    }

    New-ScriptResultsFolder
    $groupMembers | Export-Excel -Path $groupPath -WorksheetName "Group Members" -AutoSize
    Write-Host "Group members (Users only) exported to: $groupPath" -ForegroundColor Green
}

# GUI Setup
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "Inactive AD Object Management"
$form.Size = New-Object System.Drawing.Size(650, 600)
$form.StartPosition = "CenterScreen"
$form.Topmost = $true

$title = New-Object System.Windows.Forms.Label
$title.Text = "Inactive AD Object Manager ($scriptVersion)"
$title.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$title.Size = New-Object System.Drawing.Size(600, 30)
$title.TextAlign = 'MiddleCenter'
$title.Location = New-Object System.Drawing.Point(25, 20)
$form.Controls.Add($title)

# Status Label at the bottom
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Waiting..."
$statusLabel.Size = New-Object System.Drawing.Size(600, 30)
$statusLabel.Location = New-Object System.Drawing.Point(25, 450)
$statusLabel.TextAlign = 'MiddleCenter'
$statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Italic)
$form.Controls.Add($statusLabel)

$thresholdLabel = New-Object System.Windows.Forms.Label
$thresholdLabel.Text = "Inactivity Threshold (Days): $thresholdDays"
$thresholdLabel.Size = New-Object System.Drawing.Size(600, 30)
$thresholdLabel.Location = New-Object System.Drawing.Point(25, 490)
$thresholdLabel.TextAlign = 'MiddleCenter'
$thresholdLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Italic)
$form.Controls.Add($thresholdLabel)

$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Text = "Disabled Description: $disabledDescription"
$descriptionLabel.Size = New-Object System.Drawing.Size(600, 30)
$descriptionLabel.Location = New-Object System.Drawing.Point(25, 520)
$descriptionLabel.TextAlign = 'MiddleCenter'
$descriptionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Italic)
$form.Controls.Add($descriptionLabel)

# Function to set status text
function Set-Status {
    param ([string]$text)
    $statusLabel.Text = $text
    $statusLabel.Refresh()
}

# Actions + Status Wrappers
$actions = @(
    @{ Text = "1. Scrape Inactive PCs"; Action = { Set-Status "Scraping inactive PCs..."; GetPCs; Set-Status "Waiting..." } },
    @{ Text = "2. Add Description to PCs"; Action = { Set-Status "Adding description to PCs..."; Add-Description-To-PCs; Set-Status "Waiting..." } },
    @{ Text = "3. Disable Scraped PCs"; Action = { Set-Status "Disabling PCs..."; Disable-PCs; Set-Status "Waiting..." } },
    @{ Text = "4. Scrape Inactive Users"; Action = { Set-Status "Scraping inactive users..."; GetUsers; Set-Status "Waiting..." } },
    @{ Text = "5. Add Description to Users"; Action = { Set-Status "Adding description to users..."; Add-Description-To-Users; Set-Status "Waiting..." } },
    @{ Text = "6. Disable Scraped Users"; Action = { Set-Status "Disabling users..."; Disable-Users; Set-Status "Waiting..." } },
    @{ Text = "7. Set Disabled Description"; Action = { Set-Status "Waiting..."; Set-Description } },
    @{ Text = "8. Set Inactivity Threshold"; Action = { Set-Status "Waiting..."; Set-Threshold } },
    @{ Text = "9. Scrape Groups"; Action = { Set-Status "Scraping groups..."; GetGroups; Set-Status "Waiting..." } }
)

# Center buttons horizontally
$buttonWidth = 380
$centerX = ($form.ClientSize.Width - $buttonWidth) / 2
$y = 70
foreach ($action in $actions) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $action.Text
    $btn.Size = New-Object System.Drawing.Size($buttonWidth, 30)
    $btn.Location = New-Object System.Drawing.Point($centerX, $y)
    $btn.Add_Click($action.Action)
    $form.Controls.Add($btn)
    $y += 40
}

$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()