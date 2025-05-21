# Export-SPOCustomerTemplate.ps1
# --------------------------------
# Extracts a PnP site template (.pnp) from a source SharePoint site and saves it for reuse

# Configurable variables
$sourceSiteUrl = "https://twistech.sharepoint.com/sites/Initiamed"  # Change this to your template site
$templatePath = "C:\Twistech\Templates\CustomerSiteTemplate.pnp"          # Desired output location

# Ensure output folder exists
$folder = Split-Path $templatePath
if (-not (Test-Path $folder)) {
    New-Item -ItemType Directory -Path $folder | Out-Null
}

# Load PnP PowerShell (Install if needed)
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Install-Module -Name "PnP.PowerShell" -Force -Scope CurrentUser
}
Import-Module PnP.PowerShell

# Connect to the site
try {
    Write-Host "Connecting to $sourceSiteUrl..." -ForegroundColor Cyan
    Connect-PnPOnline -Url $sourceSiteUrl
    Write-Host "Connected successfully." -ForegroundColor Green
} catch {
    Write-Error "Failed to connect to SharePoint: $_"
    exit 1
}

# Export the site template
try {
    Write-Host "Exporting site template..." -ForegroundColor Yellow
    Get-PnPSiteTemplate -Out $templatePath -Handlers Lists, Branding, Navigation, Pages, ContentTypes, Fields
    Write-Host "Template saved to $templatePath" -ForegroundColor Green
} catch {
    Write-Error "Failed to export site template: $_"
}
