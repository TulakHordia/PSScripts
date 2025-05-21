# Requires the ImportExcel module
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$savePath = "C:\Twistech\Script Results"

function New-ScriptResultsFolder {
    if (-not (Test-Path -Path $savePath)) {
        New-Item -Path $savePath -ItemType Directory | Out-Null
    }
}

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Folder Permissions Exporter"
$form.Size = New-Object System.Drawing.Size(500, 260)
$form.StartPosition = "CenterScreen"

# Label
$label = New-Object System.Windows.Forms.Label
$label.Text = "Select folder to scan permissions:"
$label.Location = New-Object System.Drawing.Point(20, 20)
$label.Size = New-Object System.Drawing.Size(300, 20)
$form.Controls.Add($label)

# TextBox
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(20, 50)
$textBox.Size = New-Object System.Drawing.Size(340, 20)
$form.Controls.Add($textBox)

# Browse Button
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Browse..."
$browseButton.Location = New-Object System.Drawing.Point(370, 47)
$browseButton.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderBrowser.ShowDialog() -eq "OK") {
        $textBox.Text = $folderBrowser.SelectedPath
    }
})
$form.Controls.Add($browseButton)

# Export Button
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Text = "Export Permissions to Excel"
$exportButton.Location = New-Object System.Drawing.Point(20, 90)
$exportButton.Size = New-Object System.Drawing.Size(440, 40)
$exportButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$exportButton.BackColor = [System.Drawing.Color]::LightSteelBlue

$exportButton.Add_Click({
    $folderPath = $textBox.Text
    if (-not (Test-Path $folderPath)) {
        [System.Windows.Forms.MessageBox]::Show("Invalid folder path!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $results = @()

    Get-ChildItem -Path $folderPath -Recurse -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $acl = Get-Acl $_.FullName
        foreach ($access in $acl.Access) {
            $results += [PSCustomObject]@{
                FolderPath       = $_.FullName
                IdentityReference = $access.IdentityReference
                FileSystemRights = $access.FileSystemRights
                AccessControlType = $access.AccessControlType
                IsInherited       = $access.IsInherited
            }
        }
    }

    # Also include root folder itself
    $rootAcl = Get-Acl $folderPath
    foreach ($access in $rootAcl.Access) {
        $results += [PSCustomObject]@{
            FolderPath       = $folderPath
            IdentityReference = $access.IdentityReference
            FileSystemRights = $access.FileSystemRights
            AccessControlType = $access.AccessControlType
            IsInherited       = $access.IsInherited
        }
    }

    # Save to Excel
    New-ScriptResultsFolder
    $outputPath = "$savePath\Folder_Permissions_$(Get-Date -Format 'yyyyMMdd_HHmm').xlsx"
    $results | Export-Excel -Path $outputPath -AutoSize -Title "Folder Permissions" -FreezeTopRow -BoldTopRow

    [System.Windows.Forms.MessageBox]::Show("Export completed:`n$outputPath", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

$form.Controls.Add($exportButton)

# Run the form
[void]$form.ShowDialog()