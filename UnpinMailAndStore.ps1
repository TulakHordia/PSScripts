# Unpin Mail
$apps = (New-Object -ComObject shell.application).namespace((Get-Item "$Env:ProgramFiles\WindowsApps").fullname).Items()
$apps | Where-Object { $_.name -eq "Mail" } | ForEach-Object { $_.InvokeVerb("taskbarunpin") }

# Unpin Store
$apps | Where-Object { $_.name -eq "Microsoft Store" } | ForEach-Object { $_.InvokeVerb("taskbarunpin") }
