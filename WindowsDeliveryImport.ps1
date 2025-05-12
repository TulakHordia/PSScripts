# Optimized one-time GPO import script
Import-Module GroupPolicy

# Configuration
$backupPath = 'C:\Twistech\GPO'  # Path where the backup is stored
$backupId = 'C78CC909-844D-48D4-9A7F-364F7B5C3C3D' # Your Backup ID
$gpoName = 'Windows Update Delivery Optimization' # Name for the imported GPO

try {
    # Check if GPO already exists
    $existingGPO = Get-GPO -Name $gpoName -ErrorAction SilentlyContinue

    if ($existingGPO) {
        Write-Host "GPO '$gpoName' already exists in domain '$((Get-ADDomain).DNSRoot)'." -ForegroundColor Yellow
    }
    else {
        # Import the GPO
        $params = @{
            BackupId       = $backupId
            TargetName     = $gpoName
            Path           = $backupPath
            CreateIfNeeded = $true
        }
        Import-GPO @params

        Write-Host "GPO '$gpoName' imported successfully into domain '$((Get-ADDomain).DNSRoot)'." -ForegroundColor Green
    }
}
catch {
    Write-Host "An error occurred while importing the GPO: $_" -ForegroundColor Red
}
