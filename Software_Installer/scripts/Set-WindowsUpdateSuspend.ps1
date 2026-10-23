$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

$serviceName = "wuauserv"
$regKey = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
$arrRegVal = @()
$arrRegVal = $arrRegVal + [pscustomobject]@{
    RegName = "DoNotConnectToWindowsUpdateInternetLocations";
    RegVal  = 1
}

if (!(Test-Path -Path $regKey -ErrorAction SilentlyContinue)) {
    New-Item -Path $regKey -Force
}

Write-Output "Setting 30 day Windows Update deferral$($regKey)..."

foreach ($regVal in $arrRegVal) {
    Write-Output "Setting $($regVal.RegName) to value of $($regVal.RegVal) under $($regKey)"
    Set-ItemProperty -Path $regKey -Name $regVal.regName -Value $regVal.RegVal
}

Write-Output "Restarting $($serviceName)"
Restart-Service -Name $serviceName
