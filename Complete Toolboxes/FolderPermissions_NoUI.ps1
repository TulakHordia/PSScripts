# Requires ImportExcel module
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel

# === CONFIGURATION ===
$RootFolder = Read-Host "Enter the full path to the root folder"
$MaxDepth = 3  # Set the recursion depth here
$savePath = "C:\Twistech\Script Results"

# === Create Output Directory If Needed ===
if (-not (Test-Path -Path $savePath)) {
    New-Item -Path $savePath -ItemType Directory | Out-Null
}

# === Function to Calculate Folder Depth ===
function Get-FolderDepth {
    param ($basePath, $currentPath)
    ($currentPath -replace [regex]::Escape($basePath), '') -split '\\' | Where-Object { $_ -ne '' } | Measure-Object | Select-Object -ExpandProperty Count
}

# === Collect Folder Permissions up to Specified Depth ===
$results = @()

if (-not (Test-Path $RootFolder)) {
    Write-Error "The path '$RootFolder' does not exist."
    exit
}

# Include the root folder itself
try {
    $rootAcl = Get-Acl -Path $RootFolder
    foreach ($access in $rootAcl.Access) {
        $results += [PSCustomObject]@{
            FolderPath        = $RootFolder
            IdentityReference = $access.IdentityReference
            FileSystemRights  = $access.FileSystemRights
            AccessControlType = $access.AccessControlType
            IsInherited       = $access.IsInherited
        }
    }
} catch {
    Write-Warning "Failed to get ACL for $RootFolder"
}

# Recursively scan directories up to the desired depth
Get-ChildItem -Path $RootFolder -Directory -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
    $depth = Get-FolderDepth -basePath $RootFolder -currentPath $_.FullName
    if ($depth -le $MaxDepth) {
        try {
            $acl = Get-Acl $_.FullName
            foreach ($access in $acl.Access) {
                $results += [PSCustomObject]@{
                    FolderPath        = $_.FullName
                    IdentityReference = $access.IdentityReference
                    FileSystemRights  = $access.FileSystemRights
                    AccessControlType = $access.AccessControlType
                    IsInherited       = $access.IsInherited
                }
            }
        } catch {
            Write-Warning "Failed to get ACL for $($_.FullName)"
        }
    }
}

# === Export to Excel ===
$outputPath = "$savePath\Folder_Permissions_$(Get-Date -Format 'yyyyMMdd_HHmm').xlsx"
$results | Export-Excel -Path $outputPath -AutoSize -Title "Folder Permissions" -FreezeTopRow -BoldTopRow

Write-Host "`n✅ Export complete: $outputPath" -ForegroundColor Green