[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$TaskName = 'AVD UAT Automated Checksv1',
    [string]$ScriptPath,
    [string]$InputFile,
    [string]$ConfigurationFile,
    [string]$OutputDirectory = 'C:\Program Files\ArmyAVD\UAT',
    [string]$UatDirectory = 'C:\Program Files\ArmyAVD\UAT',
    [ValidateSet('PowerShell', 'Executable')]
    [string]$ExecutionMode = 'PowerShell',
    [ValidateSet('AtStartup', 'Daily')]
    [string]$TriggerType = 'AtStartup',
    [ValidatePattern('^(?:[01]\d|2[0-3]):[0-5]\d$')]
    [string]$DailyTime = '03:00'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $ScriptPath) {
    $ScriptPath = if ($ExecutionMode -eq 'Executable') { Join-Path $PSScriptRoot 'AVD-UAT-Checklist-Automated.exe' } else { Join-Path $PSScriptRoot 'AVD-UAT-Checklist-Automated.ps1' }
}
if (-not $ConfigurationFile) {
    $defaultConfigurationFile = Join-Path $PSScriptRoot 'avd-uat.config.json'
    if (Test-Path -LiteralPath $defaultConfigurationFile -PathType Leaf) { $ConfigurationFile = $defaultConfigurationFile }
}
if (-not $InputFile) {
    $InputFile = Get-ChildItem -Path (Join-Path $UatDirectory 'AVD Test Plan*') -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike '~$*' -and $_.Extension -in @('.xlsx', '.xlsm') } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1 -ExpandProperty FullName
}
elseif (-not (Test-Path -LiteralPath $InputFile -PathType Leaf)) {
    $pattern = if ([IO.Path]::IsPathRooted($InputFile)) { $InputFile } else { Join-Path $UatDirectory $InputFile }
    $InputFile = Get-ChildItem -Path $pattern -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike '~$*' -and $_.Extension -in @('.xlsx', '.xlsm') } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1 -ExpandProperty FullName
}
if (-not (Test-Path -LiteralPath $ScriptPath)) { throw "Runner not found: $ScriptPath" }
if (-not $InputFile -or -not (Test-Path -LiteralPath $InputFile)) { throw "No matching AVD Test Plan .xlsx or .xlsm checklist was found." }
if ($ConfigurationFile -and -not (Test-Path -LiteralPath $ConfigurationFile -PathType Leaf)) { throw "Configuration file not found: $ConfigurationFile" }

$ScriptPath = (Resolve-Path -LiteralPath $ScriptPath).Path
$InputFile = (Resolve-Path -LiteralPath $InputFile).Path
if ($ConfigurationFile) { $ConfigurationFile = (Resolve-Path -LiteralPath $ConfigurationFile).Path }
$OutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
$executable = if ($ExecutionMode -eq 'Executable') { $ScriptPath } else { Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe' }
$arguments = if ($ExecutionMode -eq 'Executable') {
    '-InputFile "{0}" -OutputDirectory "{1}"' -f $InputFile, $OutputDirectory
}
else {
    '-NoProfile -NonInteractive -File "{0}" -InputFile "{1}" -OutputDirectory "{2}"' -f $ScriptPath, $InputFile, $OutputDirectory
}
if ($ConfigurationFile) { $arguments += ' -ConfigurationFile "{0}"' -f $ConfigurationFile }
$arguments += ' -UatDirectory "{0}"' -f $UatDirectory
$action = New-ScheduledTaskAction -Execute $executable -Argument $arguments
$trigger = if ($TriggerType -eq 'AtStartup') {
    $startupTrigger = New-ScheduledTaskTrigger -AtStartup
    $startupTrigger.Delay = 'PT5M'
    $startupTrigger
}
else {
    New-ScheduledTaskTrigger -Daily -At $DailyTime
}
$principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 30) -MultipleInstances IgnoreNew

if ($PSCmdlet.ShouldProcess($TaskName, 'Register SYSTEM scheduled task')) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description 'Runs unattended AVD UAT machine checks and creates HTML, CSV, and XLSX reports.' -Force | Out-Null
    Write-Output "Scheduled task registered: $TaskName"
    if ($TriggerType -eq 'AtStartup') { Write-Output 'Startup delay: 5 minutes after boot to allow services to initialize; all prerequisites must still pass.' }
    Write-Output "Report directory: $OutputDirectory"
}


# SIG # Begin signature block
# MIIHMAYJKoZIhvcNAQcCoIIHITCCBx0CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAt5D/y236TUmik
# /M/rkEXXknL83yYTwwkrn2sPPit7U6CCBBwwggQYMIICgKADAgECAhARPahbg4yR
# u0x0RWe84V06MA0GCSqGSIb3DQEBCwUAMCQxIjAgBgNVBAMMGUFWRCBVQVQgVGVz
# dCBDb2RlIFNpZ25pbmcwHhcNMjYwOTIzMTk0MjUxWhcNMzEwOTIzMTk1MjUwWjAk
# MSIwIAYDVQQDDBlBVkQgVUFUIFRlc3QgQ29kZSBTaWduaW5nMIIBojANBgkqhkiG
# 9w0BAQEFAAOCAY8AMIIBigKCAYEAr3U0akZAjTYFpKBuGxG/XOZrFq1TiCKvagLt
# k5EbZnFvlNami9JqgzikK4A/wFJxQBdhaffseFIxS0T7HearbpAtDuQTVdkhcC03
# uBMs68EyzlggBzT2hOH7v9ebpQ3l0iz23ubNUXAtevgda1KVHfpNdTo1I1kSE6vt
# dWzVx4XyrJDC/2kK7yMpFZAw1B0VqrNOt2EQDrBAYofgIQxB6QEmp7b7O33qlIBq
# hGJFziEsTqrGbr0oCQKiXblwmUS1qKl+hMb93xkZ29WiC3PCOpnW+jwEN5zfGRqH
# 1Q0j1eVEoFTNWuEJ81N7HTi7tNAC71gO/4PNQJ0562UFkovR9fUX00UUBjKCBSqy
# nqYIKKqSErrOAFYepr4bOpheVPNsWxxwjxLWfBze92kX9IgNN3otMqTvVFY2nYXl
# hhbN217V3yGoqVmaA6a+AIiBYhXt3WegcDtP02MX/mZvbZpeHC6QpqlKHZVxgdth
# cN/5evLCaoGxxBO85VJwFM/JKbUpAgMBAAGjRjBEMA4GA1UdDwEB/wQEAwIHgDAT
# BgNVHSUEDDAKBggrBgEFBQcDAzAdBgNVHQ4EFgQUkggBwTJ4XHabRzNkW8BAJOxx
# UXAwDQYJKoZIhvcNAQELBQADggGBAAZBilHu50OctPP6mSsw/AxFk/gdNzSr/1dC
# Hi24+oU6t6bcZcRp1JJXKbipWLHfR1z3QMjuOzbwVEOlBfq0rQRygVVrKkH21MXs
# ubIjl1AWXdkwrnbJl0TpiCBDQBBP+obPS+eJY524ovOBY2tZvI3jgOxIVJDDMRwu
# hB/5BzR/n7V1VUFRzRPeVWTsgjWC+qbkbzrPMgqk1yOinBtFdNRiU67lpMR8OKZx
# UiMVhacc/u8BsEaqFVLJD2jcXe8tnjqYpFHUdb1J6pe0rooMf6oTbQtBtokM5hcx
# rh+kkvuQeY0V340WFGKw1DfTJE9HCnwkGPR94KgVauYWmF3dF3qDlvAT6mUqZxT2
# LyR98EnVSlbS+cFWf/2jSEGAZlbv1axCS1JHaypwF3Lo0JUGbzkgE4eupgv65w0L
# bzOkrUw/s3ZTVzuzfHW4KFiGOhzwqO26lsviuWZtefpsPB7mhp6B65wo7Bt+wTrx
# Nbarmp4AGEZFrhszz4hGkaX1swYgaDGCAmowggJmAgEBMDgwJDEiMCAGA1UEAwwZ
# QVZEIFVBVCBUZXN0IENvZGUgU2lnbmluZwIQET2oW4OMkbtMdEVnvOFdOjANBglg
# hkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3
# DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEV
# MC8GCSqGSIb3DQEJBDEiBCBqVY8z3i94PpQ5t6RLdvOgIb4cVvlgj2wYmFaMjLDb
# TTANBgkqhkiG9w0BAQEFAASCAYAPRAa2IVyDg08iDqj5aCbPipEIHpcRsuzmwvyN
# m7Ht+jB3vZJjzN8AdbBmIntwpMrAJSw1fvtnt5F6ogX2Ohx/3qcTlWd89aAULe64
# +WqlKOm3pot7jIzy3fGP95x3nsxqtRPzdTQ2S/07fSay8gBTLiLQPsi5I1bthOXA
# uRwKYPQL0QZZs5vxzz706h1UAND2nzolCoGCngYdDdXTp2q5pkxbH4bgDO82hCg1
# JgnUzUz/AOHbrRTOG0Uiehzb5Cyf+nu/t/f6W/rH+JWAKWrK95fvA8Rd8k8eSNDT
# Mnzpo6cPgsqQeZf8FiG4o9a/ojJw/xy9zSU9DHr1G6N0s6suq0J6c+8i6B0LmSLS
# Oc/rzKSjgy90vnbQvhGljuWRAKguMppOjTruSg4FyolB7TSzjf1RvEjRRuyA+BJD
# Mf3f1CZ/0K58iuqbUGpIZayp1XsR38JH5RjjEo+yLq95vyFDWcnDX05WAKmsM6iK
# E63mHkC8c+8HqFLMV0yyb0y85lI=
# SIG # End signature block
