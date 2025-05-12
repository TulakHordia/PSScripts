$SentinelSiteToken = "input_token";
$tempPath = "C:\temp\Tools";
$SentinelPath = "$tempPath\S1.exe";

function TempPath
{
	if(Test-Path $tempPath)
	{
		Remove-Item "$tempPath\*";
	}
	else
	{
		New-item $tempPath -ItemType Directory;
	}
}

function SentinelInstall
{
	if((Get-WmiObject win32_operatingsystem | Select-Object  osarchitecture).osarchitecture -eq "64-bit")
	{
		Set-Location $tempPath;
		Start-Process S1.exe -ArgumentList "/SILENT /SITE_TOKEN=$SentinelSiteToken"
	}
	else {
		Read-Host "Press enter to exit.."
	}
	
TempPath;
SentinelInstall;
