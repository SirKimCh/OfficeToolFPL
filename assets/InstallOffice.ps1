# Create and set the working folder
$workingFolder = "C:\OfficeSetup"
Write-Host "Creating and setting the working folder at $workingFolder..."
if (-Not (Test-Path -Path $workingFolder)) {
    New-Item -ItemType Directory -Path $workingFolder
}
Set-Location -Path $workingFolder

# Download the file from the provided link
$setupUrl = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
$setupFile = "setup.exe"
Write-Host "Downloading setup.exe from $setupUrl..."
Invoke-WebRequest -Uri $setupUrl -OutFile $setupFile

# Create the Configuration.xml file and write the provided contents
Write-Host "Writing Configuration.xml..."
$configurationXml = @"
<Configuration ID="7c6799fa-c5fb-403f-a475-a64768c92259">
  <Add OfficeClientEdition="64" Channel="Current" MigrateArch="TRUE">
    <Product ID="O365ProPlusEEANoTeamsRetail">
      <Language ID="MatchOS" />
      <Language ID="vi-vn" />
      <Language ID="MatchPreviousMSI" />
      <ExcludeApp ID="Access" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="OneNote" />
      <ExcludeApp ID="Outlook" />
      <ExcludeApp ID="Publisher" />
      <ExcludeApp ID="Bing" />
    </Product>
    <Product ID="ProofingTools">
      <Language ID="vi-vn" />
    </Product>
  </Add>
  <Property Name="SharedComputerLicensing" Value="0" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="DeviceBasedLicensing" Value="0" />
  <Property Name="SCLCacheOverride" Value="0" />
  <Updates Enabled="TRUE" />
  <RemoveMSI />
  <AppSettings>
    <User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" />
    <User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" />
    <User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas" />
  </AppSettings>
  <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>
"@
$configurationXml | Out-File -FilePath "Configuration.xml" -Encoding UTF8

# Run the setup command
Write-Host "Running setup.exe with Configuration.xml..."
Start-Process -FilePath ".\setup.exe" -ArgumentList "/configure Configuration.xml" -Wait

# Run the activate command
Write-Host "Running KMS activation and renewal task..."
& ([ScriptBlock]::Create((irm https://get.activated.win))) /KMS-ActAndRenewalTask

Write-Host "Done!"
