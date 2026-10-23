$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

function Get-TimeStamp { return "[{0:MM/dd/yy} {0:HH:mm:ss} UTC]" -f (Get-Date).ToUniversalTime() }

############################
## Stage ECMA Certificate ##
############################

Write-Output "$(Get-TimeStamp) - Starting Script - Stage ECMA Certificate"

$TmpDirectoryPath = "C:\Temp"
$BootstrapDirectoryPath = Join-Path -Path $TmpDirectoryPath -ChildPath "bootstrap"
if (-not (Test-Path $BootstrapDirectoryPath)) {
    New-Item -Path $BootstrapDirectoryPath -ItemType Directory -Force | Out-Null
}

$DownloadedCertificatesDirectoryPath = "C:\downloaded-certs"

$EcmaCertificateFilePath = Join-Path -Path $BootstrapDirectoryPath -ChildPath "ECMA-WCF-TRUSTED.crt"
if (-not (Test-Path $EcmaCertificateFilePath)) {
    throw "Failed to find ECMA certificate '$EcmaCertificateFilePath'"
}

$EcmaCertificateDestinationDirectory = Join-Path -Path $DownloadedCertificatesDirectoryPath -ChildPath "root"
New-Item -Path $EcmaCertificateDestinationDirectory -ItemType Directory -Force | Out-Null

Write-Output "$(Get-TimeStamp) - Moving ECMA certificate to directory '$DownloadedCertificatesDirectoryPath' for later installation"
$EcmaCertificateDestinationPath = Join-Path -Path $EcmaCertificateDestinationDirectory -ChildPath "ECMA-WCF-TRUSTED.crt"
Move-Item -Path $EcmaCertificateFilePath -Destination $EcmaCertificateDestinationPath -Force
if (-not (Test-Path $EcmaCertificateDestinationPath)) {
    throw "Failed to find ECMA certificate after move '$EcmaCertificateDestinationPath'"
}

Write-Output "$(Get-TimeStamp) - Completed Script - Stage ECMA Certificate"