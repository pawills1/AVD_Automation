$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

function Get-TimeStamp { return "[{0:MM/dd/yy} {0:HH:mm:ss} UTC]" -f (Get-Date).ToUniversalTime() }

#################################
## Cleanup Working Directories ##
#################################

Write-Output "$(Get-TimeStamp) - Starting Script - Cleanup Working Directories"

Write-Output "$(Get-TimeStamp) - Cleaning up working directories"
Remove-Item -Path "C:\AVDOffice.xml" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "C:\downloaded-certs" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "C:\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "C:\virtual-desktop-optimization-tool" -Recurse -Force -ErrorAction SilentlyContinue
&attrib.exe +h +s "C:\Program Files\AVDAppInstallLogs"

Write-Output "$(Get-TimeStamp) - Completed Script - Cleanup Working Directories"
