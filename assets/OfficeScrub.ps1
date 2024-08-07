# Set the working folder to the OfficeScrub folder in the user's AppData
$workingFolder = "$env:APPDATA\OfficeScrub"
Write-Output "Setting the working folder to: $workingFolder"

if (-not (Test-Path -Path $workingFolder)) {
    New-Item -ItemType Directory -Path $workingFolder -Force
    Write-Output "Created the working folder."
} else {
    Write-Output "The working folder already exists."
}

Set-Location -Path $workingFolder
Write-Output "Changed the working directory to: $workingFolder"

# Download the file from the URL
$url = "https://github.com/abbodi1406/WHD/raw/master/scripts/OfficeScrubber_12r.zip"
$outputFile = "$workingFolder\OfficeScrup.zip"
Write-Output "Downloading the file from $url to $outputFile"

Invoke-WebRequest -Uri $url -OutFile $outputFile
Write-Output "Download complete."

# Unzip the file to the working folder
Write-Output "Unzipping the file $outputFile to $workingFolder"
Expand-Archive -Path $outputFile -DestinationPath $workingFolder -Force
Write-Output "Unzipping complete."

# Run the command OfficeScrubber.cmd /A
$cmdPath = "$workingFolder\OfficeScrubber.cmd"
Write-Output "Running the command $cmdPath /A"

Start-Process -FilePath $cmdPath -ArgumentList "/A" -NoNewWindow -Wait
Write-Output "Command execution complete."
