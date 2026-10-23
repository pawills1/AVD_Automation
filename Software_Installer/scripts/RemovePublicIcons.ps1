$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

function Get-TimeStamp { return "[{0:MM/dd/yy} {0:HH:mm:ss} UTC]" -f (Get-Date).ToUniversalTime() }

#####################################
## Remove All Public Desktop Icons ##
#####################################

Write-Output "$(Get-TimeStamp) - Starting Script - Remove All Public Desktop Icons"

Write-Output "$(Get-TimeStamp) - Removing all icons and shortcuts from public desktop"
Remove-Item -Path "C:\Users\Public\Desktop\*" -Force -ErrorAction SilentlyContinue

Write-Output "$(Get-TimeStamp) - Completed Script - Remove All Public Desktop Icons"
