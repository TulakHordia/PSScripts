# Define paths
$ODTPath = "C:\ODT"
$ConfigPath = "$ODTPath\config.xml"
$SetupExe = "$ODTPath\setup.exe"

# Create folder
New-Item -ItemType Directory -Path $ODTPath -Force

# Download ODT
Invoke-WebRequest -Uri "https://aka.ms/office365odt" -OutFile $SetupExe

# Create config.xml with Shared Computer Activation enabled
$configXml = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="MonthlyEnterprise">
    <Product ID="O365ProPlusRetail">
      <Language ID="en-us" />
      <Property Name="SharedComputerLicensing" Value="1" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="1" />
</Configuration>
"@
$configXml | Out-File -FilePath $ConfigPath -Encoding UTF8

# Install Office
Start-Process -FilePath $SetupExe -ArgumentList "/configure $ConfigPath" -Wait
