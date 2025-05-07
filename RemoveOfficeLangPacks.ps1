$OfficePath = "C:\Program Files\Microsoft Office\Office16"

# Lang packs
$LanguagePacks = Get-ChildItem -Path "$OfficePath" -Filter "*LanguagePack*" | Where-Object { $_.PSIsContainer }

if ($LanguagePacks) {
    foreach ($pack in $LanguagePacks) {
        $UninstallCommand = "$pack.FullName\Setup.exe" /uninstall /config $OfficePath\SilentUninstallConfig.xml
        Write-Host "Uninstalling language pack from $pack.FullName..."
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c", $UninstallCommand -Wait -NoNewWindow
    }
    Write-Host "All language packs have been uninstalled."
} else {
    Write-Host "No language packs found."
}
