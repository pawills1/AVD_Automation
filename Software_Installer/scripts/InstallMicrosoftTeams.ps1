$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

function Get-TimeStamp { return "[{0:MM/dd/yy} {0:HH:mm:ss} UTC]" -f (Get-Date).ToUniversalTime() }

#############################
## Install Microsoft Teams ##
#############################

Write-Output "$(Get-TimeStamp) - Starting Script - Install Microsoft Teams"

$TmpDirectoryPath = "C:\Temp"
$BootstrapDirectoryPath = Join-Path -Path $TmpDirectoryPath -ChildPath "bootstrap"
if (-not (Test-Path $BootstrapDirectoryPath)) {
    New-Item -Path $BootstrapDirectoryPath -ItemType Directory -Force | Out-Null
}

$MicrosoftTeamsRegistryItemPath = "HKLM:\SOFTWARE\Microsoft"
$MicrosoftTeamsDirectoryPath = Join-Path -Path $BootstrapDirectoryPath -ChildPath "teams"
$VisualCPlusPlusDownloadUrl = "https://aka.ms/vs/16/release/vc_redist.x64.exe"
$VisualCPlusPlusExeFileName = "vc_redist.x64.exe"
$VisualCPlusPlusExeFilePath = Join-Path -Path $MicrosoftTeamsDirectoryPath -ChildPath $VisualCPlusPlusExeFileName
$WebSocketServiceDownloadUrl = "https://aka.ms/msrdcwebrtcsvc/msi"
$WebSocketServiceMsiFileName = "webSocketSvc.msi"
$WebSocketServiceMsiFilePath = Join-Path -Path $MicrosoftTeamsDirectoryPath -ChildPath $WebSocketServiceMsiFileName
$MicrosoftTeamsDownloadUrl = "https://dod.teams.microsoft.us/downloads/desktopurl?env=production&plat=windows&arch=x64&managedInstaller=true&download=true"
$MicrosoftTeamsMsiFileName = "teams.msi"
$MicrosoftTeamsMsiFilePath = Join-Path -Path $MicrosoftTeamsDirectoryPath -ChildPath $MicrosoftTeamsMsiFileName

Write-Output "$(Get-TimeStamp) - Configuring registry for Microsoft Teams AVD Media Optimization"
New-Item -Path $MicrosoftTeamsRegistryItemPath -Name "Teams" -Force
New-ItemProperty `
    -Path "${MicrosoftTeamsRegistryItemPath}\Teams" `
    -Name "IsWVDEnvironment" `
    -PropertyType "Dword" `
    -Value 1 `
    -Force

Write-Output "$(Get-TimeStamp) - Creating working directory for Microsoft Teams '$MicrosoftTeamsDirectoryPath'"
New-Item -Path $MicrosoftTeamsDirectoryPath -ItemType Directory -Force

Write-Output "$(Get-TimeStamp) - Downloading Visual C++ for Microsoft Teams"
Invoke-WebRequest -Uri $VisualCPlusPlusDownloadUrl -OutFile $VisualCPlusPlusExeFilePath
if (-not (Test-Path $VisualCPlusPlusExeFilePath)) {
    throw "Failed to find downloaded Visual C++ file for Microsoft Teams '$VisualCPlusPlusExeFilePath'"
}

Write-Output "$(Get-TimeStamp) - Installing Visual C++ for Microsoft Teams"
Start-Process -Wait `
    -FilePath $VisualCPlusPlusExeFilePath `
    -Args "/install /quiet /norestart /log vcdist.log"

Write-Output "$(Get-TimeStamp) - Downloading WebSocket Service for Microsoft Teams"
Invoke-WebRequest -Uri $WebSocketServiceDownloadUrl -OutFile $WebSocketServiceMsiFilePath
if (-not (Test-Path $WebSocketServiceMsiFilePath)) {
    throw "Failed to find downloaded WebSocket Service file for Microsoft Teams '$WebSocketServiceMsiFilePath'"
}

Write-Output "$(Get-TimeStamp) - Installing WebSocket Service for Microsoft Teams"
Start-Process  -Wait `
    -FilePath msiexec.exe `
    -Args "/I $WebSocketServiceMsiFilePath /quiet /norestart /log webSocket.log"

Write-Output "$(Get-TimeStamp) - Downloading Microsoft Teams"
Invoke-WebRequest -Uri $MicrosoftTeamsDownloadUrl -OutFile $MicrosoftTeamsMsiFilePath
if (-not (Test-Path $MicrosoftTeamsMsiFilePath)) {
    throw "Failed to find downloaded Microsoft Teams file '$MicrosoftTeamsMsiFilePath'"
}

Write-Output "$(Get-TimeStamp) - Installing Microsoft Teams"
Start-Process  -Wait `
    -FilePath msiexec.exe `
    -Args "/I $MicrosoftTeamsMsiFilePath /quiet /norestart /log teams.log ALLUSER=1 ALLUSERS=1"

Write-Output "$(Get-TimeStamp) - Completed Script - Install Microsoft Teams"
