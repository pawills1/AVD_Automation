$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

function Get-TimeStamp { return "[{0:MM/dd/yy} {0:HH:mm:ss} UTC]" -f (Get-Date).ToUniversalTime() }

######################
## Stage ADMX Files ##
######################

Write-Output "$(Get-TimeStamp) - Starting Script - Stage ADMX"

$TmpDirectoryPath = "C:\Temp"
$BootstrapDirectoryPath = Join-Path -Path $TmpDirectoryPath -ChildPath "bootstrap"
if (-not (Test-Path $BootstrapDirectoryPath)) {
    New-Item -Path $BootstrapDirectoryPath -ItemType Directory -Force | Out-Null
}

$AdmxPacZipFilePath = Join-Path -Path $BootstrapDirectoryPath -ChildPath "admxpac.zip"
if (-not (Test-Path $AdmxPacZipFilePath)) {
    throw "Failed to find downloaded ADMX zip file '$AdmxPacZipFilePath'"
}

Write-Output "$(Get-TimeStamp) - Extracting ADMX zip file to directory 'C:\windows\PolicyDefinitions'"
Expand-Archive -Path $AdmxPacZipFilePath -DestinationPath "C:\windows\PolicyDefinitions" -Force

Write-Output "$(Get-TimeStamp) - Completed Script - Stage ADMX"
