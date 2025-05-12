$RegistryPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Search'
$PropertyName = 'SearchboxTaskbarMode'
$NewValue = '1'

Set-ItemProperty -Path $RegistryPath -Name $PropertyName -Value $NewValue