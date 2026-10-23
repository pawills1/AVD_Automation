$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

function Get-TimeStamp { return "[{0:MM/dd/yy} {0:HH:mm:ss} UTC]" -f (Get-Date).ToUniversalTime() }

##################################
## Check Windows Update Version ##
##################################

Write-Output "$(Get-TimeStamp) - Starting Script - Check Windows Update Version"

$KBUpdate = (Get-HotFix).where({ $PSItem.HotFixID -eq $env:KBVersionNumber })
if ($KBUpdate.Count -gt 0) {
    $KBUpdate
    Write-Output "$(Get-TimeStamp) - Found expected update '$env:KBVersionNumber'"
} else {
    throw "Failed to find expected update '$env:KBVersionNumber'"
}

Write-Output "$(Get-TimeStamp) - Completed Script - Check Windows Update Version"
