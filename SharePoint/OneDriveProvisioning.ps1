# Load necessary assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Define form
$form = New-Object System.Windows.Forms.Form
$form.Text = "OneDrive Provisioning Tool"
$form.Size = New-Object System.Drawing.Size(500, 360)
$form.StartPosition = "CenterScreen"

# Tenant input label
$tenantLabel = New-Object System.Windows.Forms.Label
$tenantLabel.Text = "Enter tenant prefix (e.g. contoso):"
$tenantLabel.Size = New-Object System.Drawing.Size(460, 20)
$tenantLabel.Location = New-Object System.Drawing.Point(10, 20)
$form.Controls.Add($tenantLabel)

# Tenant input box
$tenantBox = New-Object System.Windows.Forms.TextBox
$tenantBox.Size = New-Object System.Drawing.Size(460, 20)
$tenantBox.Location = New-Object System.Drawing.Point(10, 45)
$form.Controls.Add($tenantBox)

# Email input label
$emailLabel = New-Object System.Windows.Forms.Label
$emailLabel.Text = "Enter one or more user emails (comma-separated):"
$emailLabel.Size = New-Object System.Drawing.Size(460, 20)
$emailLabel.Location = New-Object System.Drawing.Point(10, 75)
$form.Controls.Add($emailLabel)

# Email input box
$emailBox = New-Object System.Windows.Forms.TextBox
$emailBox.Multiline = $true
$emailBox.Size = New-Object System.Drawing.Size(460, 80)
$emailBox.Location = New-Object System.Drawing.Point(10, 100)
$form.Controls.Add($emailBox)

# Status label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Size = New-Object System.Drawing.Size(460, 50)
$statusLabel.Location = New-Object System.Drawing.Point(10, 245)
$statusLabel.ForeColor = [System.Drawing.Color]::DarkGreen
$form.Controls.Add($statusLabel)

# Function: Install Module if Needed
function Ensure-SPOInstalled {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell)) {
        $statusLabel.ForeColor = [System.Drawing.Color]::DarkOrange
        $statusLabel.Text = "Installing SharePoint Online module..."
        Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Force -AllowClobber -Scope CurrentUser
    }
    Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop
}

# Function: Trigger OneDrive creation
function Trigger-OneDrive {
    param ($tenantPrefix, $emails)

    try {
        Ensure-SPOInstalled

        $statusLabel.ForeColor = [System.Drawing.Color]::DarkCyan
        $statusLabel.Text = "Connecting to SharePoint Online..."

        $tenantAdminUrl = "https://$tenantPrefix-admin.sharepoint.com"
        Connect-SPOService

        $statusLabel.Text = "Connected. Triggering OneDrive creation..."
        Request-SPOPersonalSite -UserEmails $emails -NoWait

        $statusLabel.ForeColor = [System.Drawing.Color]::DarkGreen
        $statusLabel.Text = "✅ OneDrive provisioning triggered successfully."
    } catch {
        $statusLabel.ForeColor = [System.Drawing.Color]::Red
        $statusLabel.Text = "❌ Error: $($_.Exception.Message)"
    }
}

# Button
$provisionButton = New-Object System.Windows.Forms.Button
$provisionButton.Text = "Trigger OneDrive Creation"
$provisionButton.Size = New-Object System.Drawing.Size(220, 35)
$provisionButton.Location = New-Object System.Drawing.Point(10, 200)
$provisionButton.Add_Click({
    $tenant = $tenantBox.Text.Trim()
    $emails = $emailBox.Text.Trim()

    if (-not $tenant) {
        $statusLabel.ForeColor = [System.Drawing.Color]::Red
        $statusLabel.Text = "Please enter a tenant prefix."
        return
    }

    if (-not $emails) {
        $statusLabel.ForeColor = [System.Drawing.Color]::Red
        $statusLabel.Text = "Please enter at least one email."
        return
    }

    Trigger-OneDrive -tenantPrefix $tenant -emails $emails
})
$form.Controls.Add($provisionButton)

# Show form
[void]$form.ShowDialog()
