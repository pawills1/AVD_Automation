$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

function Get-TimeStamp { return "[{0:MM/dd/yy} {0:HH:mm:ss} UTC]" -f (Get-Date).ToUniversalTime() }

#####################
## Install FSLogix ##
#####################

Write-Output "$(Get-TimeStamp) - Starting Script - Install FSLogix"

$TmpDirectoryPath = "C:\Temp"
$BootstrapDirectoryPath = Join-Path -Path $TmpDirectoryPath -ChildPath "bootstrap"
if (-not (Test-Path $BootstrapDirectoryPath)) {
    New-Item -Path $BootstrapDirectoryPath -ItemType Directory -Force | Out-Null
}

$FSLogixDirectoryPath = Join-Path -Path $BootstrapDirectoryPath -ChildPath "fslogix"
$FSLogixZipFilePath = "${FSLogixDirectoryPath}.zip"

Write-Output "$(Get-TimeStamp) - Downloading FSLogix zip file"
Invoke-WebRequest -Uri "https://aka.ms/fslogix_download" -OutFile $FSLogixZipFilePath
if (-not (Test-Path $FSLogixZipFilePath)) {
    throw "Failed to find downloaded FSLogix zip file '$FSLogixZipFilePath'"
}

Start-Sleep -Seconds 10
Write-Output "$(Get-TimeStamp) - Extracting FSLogix zip file to directory '$FSLogixDirectoryPath'"
Expand-Archive -Path $FSLogixZipFilePath -DestinationPath $FSLogixDirectoryPath -Force

$FSLogixSetupFile = (Get-ChildItem -Path $FSLogixDirectoryPath -Recurse -File -Filter "FSLogixAppsSetup.exe").Where({ $PSItem.DirectoryName -ilike "*x64\Release*" })
if (($FSLogixSetupFile.Count -eq 0) -or ($null -eq $FSLogixSetupFile)) {
    throw "Failed to find FSLogix setup file 'FSLogixAppsSetup.exe'"
}

Write-Output "$(Get-TimeStamp) - Installing FSLogix"
$InstallLogFilePath = "C:\Windows\Logs\Software\FSLOGIX-install.log"
Start-Process `
    -FilePath $FSLogixSetupFile.FullName `
    -ArgumentList "ACCEPT_EULA=1 /quiet /norestart /log $InstallLogFilePath" `
    -Wait

Write-Output "$(Get-TimeStamp) - Dumping FSLogix installation log"
Get-Content -Path $InstallLogFilePath
Remove-Item -Path $InstallLogFilePath -Force -ErrorAction SilentlyContinue | Out-Null

Write-Output "$(Get-TimeStamp) - Dumping installed FSLogix version"
Start-Process `
    -FilePath "C:\Program Files\FSLogix\Apps\frx.exe" `
    -ArgumentList "version" `
    -NoNewWindow `
    -Wait

Write-Output "$(Get-TimeStamp) - Completed Script - Install FSLogix"
