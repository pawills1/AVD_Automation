$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

function Get-TimeStamp { return "[{0:MM/dd/yy} {0:HH:mm:ss} UTC]" -f (Get-Date).ToUniversalTime() }

##################################
## Stage ECA Root 4 Certificate ##
##################################

Write-Output "$(Get-TimeStamp) - Starting Script - Stage ECA Root 4 Certificate"

$TmpDirectoryPath = "C:\Temp"
$BootstrapDirectoryPath = Join-Path -Path $TmpDirectoryPath -ChildPath "bootstrap"
if (-not (Test-Path $BootstrapDirectoryPath)) {
    New-Item -Path $BootstrapDirectoryPath -ItemType Directory -Force | Out-Null
}

$DownloadedCertificatesDirectoryPath = "C:\downloaded-certs"

$EcaRootCertificateFilePath = Join-Path -Path $BootstrapDirectoryPath -ChildPath "ECA Root CA 4.cer"
if (-not (Test-Path $EcaRootCertificateFilePath)) {
    throw "Failed to find ECA Root 4 certificate '$EcaRootCertificateFilePath'"
}

$EcaRootCertificateDestinationDirectory = Join-Path -Path $DownloadedCertificatesDirectoryPath -ChildPath "root"
New-Item -Path $EcaRootCertificateDestinationDirectory -ItemType Directory -Force | Out-Null

Write-Output "$(Get-TimeStamp) - Moving ECA Root 4 certificate to directory '$DownloadedCertificatesDirectoryPath' for later installation"
$EcaRootCertificateDestinationPath = Join-Path -Path $EcaRootCertificateDestinationDirectory -ChildPath "ECA Root CA 4.cer"
Move-Item -Path $EcaRootCertificateFilePath -Destination $EcaRootCertificateDestinationPath -Force
if (-not (Test-Path $EcaRootCertificateDestinationPath)) {
    throw "Failed to find ECA Root 4 certificate after move '$EcaRootCertificateDestinationPath'"
}

Write-Output "$(Get-TimeStamp) - Completed Script - Stage ECA Root 4 Certificate"
