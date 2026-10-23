$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

function Get-TimeStamp { return "[{0:MM/dd/yy} {0:HH:mm:ss} UTC]" -f (Get-Date).ToUniversalTime() }

###########################################
## Run Virtual Desktop Optimization Tool ##
###########################################

Write-Output "$(Get-TimeStamp) - Starting Script - Virtual Desktop Optimization Tool"

$VdotDirectoryPath = "C:\virtual-desktop-optimization-tool\"
if (-not (Test-Path $VdotDirectoryPath)) {
    throw "Failed to find certificates temporary directory."
}

$VdotScript = Get-ChildItem -Path $VdotDirectoryPath -Recurse -File -Filter "Windows_VDOT.ps1"
if (($VdotScript.Count -eq 0) -or ($null -eq $VdotScript)) {
    throw "Failed to find virtual desktop optimization tool script 'Windows_VDOT.ps1' within directory '$VdotDirectoryPath'"
}

Write-Output "$(Get-TimeStamp) - Starting virtual desktop optimization tool"
& $VdotScript.FullName -Optimizations "All" -AdvancedOptimizations "Edge" -AcceptEULA

Write-Output "$(Get-TimeStamp) - Completed Script - Virtual Desktop Optimization Tool"
