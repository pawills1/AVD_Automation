[CmdletBinding()]
param(
    [string]$InputFile,
    [string]$OutputDirectory,
    [string]$UatDirectory = 'C:\Program Files\ArmyAVD\UAT',
    [string]$WorkbookTool,
    [string]$ConfigurationFile,
    [string]$LogPath,
    [ValidateRange(1, 3600)]
    [int]$PerformanceSampleSeconds = 60,
    [ValidateRange(1, 300)]
    [int]$PerformanceSampleIntervalSeconds = 5,
    [ValidateRange(1, 100)]
    [double]$MaximumAverageCpuPercent = 40,
    [ValidateRange(1, 100)]
    [double]$MaximumMemoryUsedPercent = 80,
    [ValidateRange(0.1, 10000)]
    [double]$MaximumDiskLatencyMilliseconds = 20
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$executionRoot = if ($PSScriptRoot) { $PSScriptRoot } else { [AppDomain]::CurrentDomain.BaseDirectory.TrimEnd('\') }
. (Join-Path $executionRoot 'AVD-UAT-Paths.ps1')
if (-not $OutputDirectory) { $OutputDirectory = $UatDirectory }
$OutputDirectory = New-UatResultsDirectory -RootDirectory $OutputDirectory
$logName = if ($LogPath) { Split-Path -Leaf $LogPath } else { "AVD UAT Automated Log - $(Get-Date -Format 'yyyyMMdd-HHmmss-fff').log" }
$LogPath = Join-Path $OutputDirectory $logName

function Write-RunLog {
    param([string]$Message, [ValidateSet('INFO', 'ERROR')] [string]$Level = 'INFO')
    $entry = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $Level, $Message
    Add-Content -LiteralPath $LogPath -Value $entry -Encoding UTF8
}

$prerequisiteChecker = Join-Path $executionRoot 'AVD-UAT-Prerequisite-Check.ps1'
if (-not (Test-Path -LiteralPath $prerequisiteChecker -PathType Leaf)) {
    Write-RunLog "Prerequisite gate error: checker not found: $prerequisiteChecker" 'ERROR'
    exit 3
}
try { $prerequisiteGate = & $prerequisiteChecker -OutputDirectory $OutputDirectory -Integrated }
catch {
    Write-RunLog "Prerequisite gate error: $($_.Exception.Message)" 'ERROR'
    exit 3
}
Write-RunLog "Prerequisite log: $($prerequisiteGate.LogPath)"
Write-RunLog "Prerequisite report: $($prerequisiteGate.ReportPath)"
if (-not $prerequisiteGate.Passed) {
    Write-RunLog 'Prerequisite gate failed. Automated UAT was not started.' 'ERROR'
    exit 2
}

try {
    $executionRoot = if ($PSScriptRoot) { $PSScriptRoot } else { [AppDomain]::CurrentDomain.BaseDirectory.TrimEnd('\') }
    Write-RunLog "Automated run started. Computer=$env:COMPUTERNAME User=$env:USERDOMAIN\$env:USERNAME ProcessId=$PID PowerShell=$($PSVersionTable.PSVersion)"
    Write-RunLog "ExecutionRoot=$executionRoot OutputDirectory=$OutputDirectory"
    $runner = Join-Path $executionRoot 'AVD-UAT-Checklist.ps1'
    if (-not (Test-Path -LiteralPath $runner -PathType Leaf)) { throw "Automated check engine not found: $runner" }

    $arguments = @{}
    foreach ($name in $PSBoundParameters.Keys) {
        if ($name -notin @('LogPath', 'Verbose')) { $arguments[$name] = $PSBoundParameters[$name] }
    }
    $arguments['OutputDirectory'] = $OutputDirectory
    $arguments['UatDirectory'] = $UatDirectory
    $arguments['UseExistingResultsDirectory'] = $true
    $arguments['Verbose'] = $true
    Write-RunLog "Launching check engine: $runner"
    & $runner @arguments *>&1 | ForEach-Object {
        $message = $_ | Out-String
        $message = $message.TrimEnd()
        if ($message) { Write-RunLog $message }
    }
    Write-RunLog 'Automated run completed successfully.'
    Write-RunLog "Troubleshooting log: $LogPath"
    exit 0
}
catch {
    Write-RunLog ("Automated run failed: {0}`r`n{1}" -f $_.Exception.Message, ($_ | Out-String)) 'ERROR'
    Write-RunLog "Troubleshooting log: $LogPath" 'ERROR'
    exit 1
}

# SIG # Begin signature block
# MIIHMAYJKoZIhvcNAQcCoIIHITCCBx0CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBK0pl3Icx76Amo
# zh47RAYPBLDuezJYEU8B93mO1sOo6aCCBBwwggQYMIICgKADAgECAhARPahbg4yR
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
# MC8GCSqGSIb3DQEJBDEiBCD9zeCJNTCMHEeHOXtejVbM7buOSihfEPQFehc8b8Pk
# 2jANBgkqhkiG9w0BAQEFAASCAYCczzepYb1M3lb+FfxTZ87SGrRvPU3go8xuwLvp
# slLN6tW+UNm7iWIINofb5FhNBZ+aRPE5kbRsJQmaXryH2Kc3IwPPnK5rQ4RmJNLv
# LIKw9K9EOvpljw0LhFB9sUKjV70LnvAv5eD+Xv0aeos3VxKER48xoo1Nyhptt5m8
# X3QyUIjz7YfcQ0WESuo72tmN1OFGuhasAKu6oVlbhbf4SChKoSovvzybm7HJhkF+
# ZJEqD1L4qwxSxMqCFq4g9Xr6QI+lyQX0RXEQ31V5zUfCngAXjEkHxE9NsbVj5vRQ
# t/RqFuXzekrsneZnyjSQPedbJrVVdGYhUYMMFb5wkYhaZjjnT+fuCcInlUY2H6Qc
# pnXcH790haXQ05bohngUAfW6LtcYhSolkLK/0bQWkWMdW6Q1dMf17odFxvJW6oY/
# ET4ipDzuw6zsLJDMptORCfPmAt6Vjss+/1vTIvGlOZgoIPrq14kAPvqK3kxtV8+0
# 1aqYTsSpwPWUfOV+46m4Im2zCV8=
# SIG # End signature block
