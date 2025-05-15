# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Ensure ImportExcel module is installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}
Import-Module ImportExcel

# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Export Distribution Group Members"
$form.Size = New-Object System.Drawing.Size(500, 300)
$form.StartPosition = "CenterScreen"
$form.Topmost = $true

# Output path label
$pathLabel = New-Object System.Windows.Forms.Label
$pathLabel.Text = "Export Path:"
$pathLabel.Location = New-Object System.Drawing.Point(20, 20)
$pathLabel.Size = New-Object System.Drawing.Size(80, 20)
$form.Controls.Add($pathLabel)

# Output path textbox (XLSX default)
$pathBox = New-Object System.Windows.Forms.TextBox
$pathBox.Text = "C:\Twistech\Script Results\DistributionGroupMembers.xlsx"
$pathBox.Location = New-Object System.Drawing.Point(110, 20)
$pathBox.Size = New-Object System.Drawing.Size(340, 20)
$form.Controls.Add($pathBox)

# Status label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Status: Waiting..."
$statusLabel.Location = New-Object System.Drawing.Point(20, 200)
$statusLabel.Size = New-Object System.Drawing.Size(450, 40)
$form.Controls.Add($statusLabel)

# Connect button
$connectBtn = New-Object System.Windows.Forms.Button
$connectBtn.Text = "Connect to 365"
$connectBtn.Location = New-Object System.Drawing.Point(110, 60)
$connectBtn.Size = New-Object System.Drawing.Size(120, 40)
$form.Controls.Add($connectBtn)

# Export button
$exportBtn = New-Object System.Windows.Forms.Button
$exportBtn.Text = "Start Export"
$exportBtn.Location = New-Object System.Drawing.Point(250, 60)
$exportBtn.Size = New-Object System.Drawing.Size(120, 40)
$exportBtn.Enabled = $false
$form.Controls.Add($exportBtn)

# Function: Connect to Exchange Online
$connectBtn.Add_Click({
    try {
        $statusLabel.Text = "Status: Connecting to Exchange Online..."
        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            Install-Module ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
        }
        Import-Module ExchangeOnlineManagement -Force
        Connect-ExchangeOnline -ShowBanner:$false
        $statusLabel.Text = "Status: Connected successfully!"
        $exportBtn.Enabled = $true
    } catch {
        $statusLabel.Text = "Status: Connection failed - $($_.Exception.Message)"
    }
})

# Function: Export Distribution Groups and Members
$exportBtn.Add_Click({
    $outputPath = $pathBox.Text
    $folder = Split-Path $outputPath
    if (-not (Test-Path $folder)) {
        New-Item -ItemType Directory -Path $folder | Out-Null
    }

    $statusLabel.Text = "Status: Exporting data..."
    $results = @()

    try {
        $groups = Get-DistributionGroup -ResultSize Unlimited
        foreach ($group in $groups) {
            $members = Get-DistributionGroupMember -Identity $group.Identity -ResultSize Unlimited
            if ($members) {
                foreach ($member in $members) {
                    $results += [PSCustomObject]@{
                        GroupName   = $group.DisplayName
                        GroupEmail  = $group.PrimarySmtpAddress
                        MemberName  = $member.Name
                        MemberEmail = $member.PrimarySmtpAddress
                        MemberType  = $member.RecipientType
                    }
                }
            } else {
                $results += [PSCustomObject]@{
                    GroupName   = $group.DisplayName
                    GroupEmail  = $group.PrimarySmtpAddress
                    MemberName  = "<No Members>"
                    MemberEmail = ""
                    MemberType  = ""
                }
            }
        }

        # Export to XLSX using ImportExcel
        $results | Export-Excel -Path $outputPath -WorksheetName "Group Members" -AutoSize -TableName "DGMembers"

        $statusLabel.Text = "Status: Export complete!"
    } catch {
        $statusLabel.Text = "Status: Error - $($_.Exception.Message)"
    }
})

# Show the form
[void]$form.ShowDialog()
