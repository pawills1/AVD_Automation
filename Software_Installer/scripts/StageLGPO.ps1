$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

function Get-TimeStamp { return "[{0:MM/dd/yy} {0:HH:mm:ss} UTC]" -f (Get-Date).ToUniversalTime() }

################
## Stage LGPO ##
################

# LGPO executable is ran during VM bootstrap script via VM extension

Write-Output "$(Get-TimeStamp) - Starting Script - Stage LGPO"

$TmpDirectoryPath = "C:\Temp"
$BootstrapDirectoryPath = Join-Path -Path $TmpDirectoryPath -ChildPath "bootstrap"
if (-not (Test-Path $BootstrapDirectoryPath)) {
    New-Item -Path $BootstrapDirectoryPath -ItemType Directory -Force | Out-Null
}

$LgpoDirectoryPath = Join-Path -Path $BootstrapDirectoryPath -ChildPath "lgpopak"
$LgpoExeFilePath = "${LgpoDirectoryPath}\lgpopak\deps\LGPO.exe"
$LgpoStigsDirectoryPath = "${LgpoDirectoryPath}\lgpopak\STIGS"

$LgpoZipFilePath = "${LgpoDirectoryPath}.zip"
if (-not (Test-Path $LgpoZipFilePath)) {
    throw "Failed to find downloaded LGPO zip file '$LgpoZipFilePath'"
}

Write-Output "$(Get-TimeStamp) - Extracting LGPO zip file to directory '$LgpoDirectoryPath'"
Expand-Archive -Path $LgpoZipFilePath -DestinationPath $LgpoDirectoryPath -Force
if (-not (Test-Path $LgpoExeFilePath)) {
    throw "Failed to find extracted LGPO exe file '$LgpoExeFilePath'"
}

if (-not (Test-Path $LgpoStigsDirectoryPath)) {
    throw "Failed to find extracted LGPO STIGs directory '$LgpoStigsDirectoryPath'"
}

$LgpoStigFiles = Get-ChildItem -Path $LgpoStigsDirectoryPath -Recurse -File
if ($LgpoStigFiles.Count -eq 0) {
    throw "Failed to find moved LGPO STIGs after move to directory '$LgpoStigsDirectoryPath'"
}

Write-Output "$(Get-TimeStamp) - Moving LGPO exe file to directory 'C:\Windows'"
Move-Item -Path $LgpoExeFilePath -Destination "C:\Windows" -Force
if (-not (Test-Path "C:\Windows\LGPO.exe")) {
    throw "Failed to find LGPO exe file after move to directory 'C:\Windows'"
}

Write-Output "$(Get-TimeStamp) - Moving LGPO STIGs directory to directory 'C:\LGPO-STIGs'"
Move-Item -Path $LgpoStigsDirectoryPath -Destination "C:\LGPO-STIGs" -Force
$LgpoStigFiles = Get-ChildItem -Path "C:\LGPO-STIGs" -Recurse -File
if ($LgpoStigFiles.Count -eq 0) {
    throw "Failed to find moved LGPO STIGs after move to directory 'C:\LGPO-STIGs'"
}

Write-Output "$(Get-TimeStamp) - Completed Script - Stage LGPO"
