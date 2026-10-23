$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

function Get-TimeStamp { return "[{0:MM/dd/yy} {0:HH:mm:ss} UTC]" -f (Get-Date).ToUniversalTime() }

#######################################
## Install Microsoft Office 365 Apps ##
#######################################

# Deployment tool download link and instructions: https://learn.microsoft.com/en-us/deployoffice/overview-office-deployment-tool

Write-Output "$(Get-TimeStamp) - Starting Script - Install Microsoft Office 365 Apps"

$TmpDirectoryPath = "C:\Temp"
$BootstrapDirectoryPath = Join-Path -Path $TmpDirectoryPath -ChildPath "bootstrap"
if (-not (Test-Path $BootstrapDirectoryPath)) {
    New-Item -Path $BootstrapDirectoryPath -ItemType Directory -Force | Out-Null
}

$O365AppDeploymentToolConfigurationFilePath = "C:\AVDOffice.xml"
if (-not (Test-Path $O365AppDeploymentToolConfigurationFilePath)) {
    throw "Failed to find Microsoft Office 365 app deployment tool configuration file '$O365AppDeploymentToolConfigurationFilePath'"
}

$O365AppsDirectoryPath = Join-Path -Path $BootstrapDirectoryPath -ChildPath "o365apps"
$O365AppsDeploymentToolDownloadUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17531-20046.exe"
$O365AppsDeploymentToolExeFileName = "officedeploymenttool.exe"
$O365AppsDeploymentToolExeFilePath = Join-Path -Path $O365AppsDirectoryPath -ChildPath $O365AppsDeploymentToolExeFileName

Write-Output "$(Get-TimeStamp) - Creating working directory for Microsoft Office 365 apps '$O365AppsDirectoryPath'"
New-Item -Path $O365AppsDirectoryPath -ItemType Directory -Force

Write-Output "$(Get-TimeStamp) - Downloading Microsoft Office 365 apps Deployment Tool"
Invoke-WebRequest -Uri $O365AppsDeploymentToolDownloadUrl -OutFile $O365AppsDeploymentToolExeFilePath
if (-not (Test-Path $O365AppsDeploymentToolExeFilePath)) {
    throw "Failed to find downloaded Microsoft Office 365 apps Deployment Tool '$O365AppsDeploymentToolExeFilePath'"
}

Write-Output "$(Get-TimeStamp) - Extracting Microsoft Office 365 apps Deployment Tool"
Start-Process `
    -FilePath $O365AppsDeploymentToolExeFilePath `
    -ArgumentList "/quiet /extract:$O365AppsDirectoryPath" `
    -Wait

$O365AppsSetupExeFile = Get-ChildItem -Path $O365AppsDirectoryPath -Recurse -File -Filter "setup.exe"
if ([String]::IsNullOrWhiteSpace($O365AppsSetupExeFile.FullName) -or (-not (Test-Path $O365AppsSetupExeFile.FullName))) {
    throw "Failed to find 'setup.exe' for Microsoft Office 365 apps installation"
}

Write-Output "$(Get-TimeStamp) - Installing Microsoft Office 365 apps"
Start-Process `
    -FilePath $O365AppsSetupExeFile.FullName `
    -Args "/configure $O365AppDeploymentToolConfigurationFilePath" `
    -Wait

Write-Output "$(Get-TimeStamp) - Completed Script - Install Microsoft Office 365 apps"
