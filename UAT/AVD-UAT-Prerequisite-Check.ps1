[CmdletBinding()]
param(
    [string]$OutputDirectory = (Join-Path $env:TEMP 'ArmyAVD\UAT-Prerequisites'),
    [switch]$Integrated
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
$computerName = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { 'UnknownVM' }
$logPath = Join-Path $OutputDirectory "AVD UAT Prerequisite Log - $timestamp.log"
$reportPath = Join-Path $OutputDirectory "AVD UAT Prerequisite Report - $computerName - $timestamp.html"

function New-PrerequisiteResult {
    param(
        [string]$Name,
        [bool]$Passed,
        [string]$Evidence,
        [string]$Remediation
    )
    @{
        Prerequisite = $Name
        Result = if ($Passed) { 'Pass' } else { 'Fail' }
        Evidence = $Evidence
        Remediation = if ($Passed) { '' } else { $Remediation }
    }
}

function Get-SystemExecutablePath {
    param([string]$Name)
    if ($env:PROCESSOR_ARCHITEW6432) {
        $sysnativePath = Join-Path $env:WINDIR "Sysnative\$Name"
        if (Test-Path -LiteralPath $sysnativePath -PathType Leaf) { return $sysnativePath }
    }
    $system32Path = Join-Path $env:WINDIR "System32\$Name"
    if (Test-Path -LiteralPath $system32Path -PathType Leaf) { return $system32Path }
    return $Name
}

function ConvertTo-HtmlText {
    param([string]$Value)
    if ($null -eq $Value) { return '' }
    return $Value -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;' -replace "'", '&#39;'
}

function Write-PrerequisiteLog {
    param([string]$Message, [ValidateSet('INFO', 'ERROR')] [string]$Level = 'INFO')
    $entry = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $Level, $Message
    Add-Content -LiteralPath $logPath -Value $entry -Encoding UTF8
}

function Test-EntraJoinStatus {
    param([string[]]$DsregOutput)
    return $null -ne ($DsregOutput | Where-Object { $_ -match '^\s*AzureAdJoined\s*:\s*YES\s*$' } | Select-Object -First 1)
}

function Get-EntraJoinResult {
    $dsregcmd = Get-SystemExecutablePath 'dsregcmd.exe'
    $output = & $dsregcmd /status 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "dsregcmd.exe /status exited with code $LASTEXITCODE. $($output -join ' ')"
    }
    if (Test-EntraJoinStatus -DsregOutput $output) {
        return New-PrerequisiteResult 'Microsoft Entra joined' $true 'dsregcmd Device State reports AzureAdJoined : YES.' ''
    }
    return New-PrerequisiteResult 'Microsoft Entra joined' $false 'dsregcmd Device State does not report AzureAdJoined : YES.' 'Join the VM to Microsoft Entra ID and rerun the prerequisite gate.'
}

function Get-IntuneEnrollmentResult {
    $enrollmentsPath = 'HKLM:\SOFTWARE\Microsoft\Enrollments'
    $enrollments = @(
        Get-ChildItem -Path $enrollmentsPath -ErrorAction SilentlyContinue |
            Where-Object { $_.PSChildName -match '^[0-9a-fA-F-]{36}$' } |
            ForEach-Object {
                $providerId = $null
                try { $providerId = Get-ItemPropertyValue -LiteralPath $_.PSPath -Name 'ProviderID' -ErrorAction Stop }
                catch { }
                if ([string]$providerId -eq 'MS DM Server') {
                    @{ Id = $_.PSChildName }
                }
            }
    )
    if ($enrollments.Count -eq 0) {
        return New-PrerequisiteResult 'Intune enrolled' $false 'No MS DM Server machine enrollment was found under HKLM:\SOFTWARE\Microsoft\Enrollments.' 'Enroll the VM in Intune and rerun the prerequisite gate.'
    }

    $schtasks = Get-SystemExecutablePath 'schtasks.exe'
    $scheduledTaskOutput = @(& $schtasks /Query /FO LIST /V 2>&1)
    if ($LASTEXITCODE -ne 0) {
        throw "schtasks.exe could not enumerate EnterpriseMgmt tasks. Exit code: $LASTEXITCODE. $($scheduledTaskOutput -join ' ')"
    }
    foreach ($enrollment in $enrollments) {
        $taskPrefix = "\Microsoft\Windows\EnterpriseMgmt\$($enrollment.Id)\"
        $taskMatches = @($scheduledTaskOutput | Where-Object { $_ -like "*$taskPrefix*" })
        if ($taskMatches.Count -gt 0) {
            return New-PrerequisiteResult 'Intune enrolled' $true "Found MS DM Server enrollment $($enrollment.Id) with matching EnterpriseMgmt scheduled task output." ''
        }
    }
    $ids = $enrollments | Select-Object -ExpandProperty Id
    return New-PrerequisiteResult 'Intune enrolled' $false "Found MS DM Server enrollment(s), but no matching EnterpriseMgmt scheduled tasks: $($ids -join ', ')." 'Repair or re-enroll the VM in Intune and rerun the prerequisite gate.'
}

function Get-RunningServiceResult {
    param(
        [string]$PrerequisiteName,
        [string]$ServiceName,
        [string]$Remediation
    )
    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($null -eq $service) {
        return New-PrerequisiteResult $PrerequisiteName $false "Service '$ServiceName' is not installed." $Remediation
    }
    if ($service.Status -ne 'Running') {
        return New-PrerequisiteResult $PrerequisiteName $false "Service '$ServiceName' is $($service.Status); expected Running." $Remediation
    }
    return New-PrerequisiteResult $PrerequisiteName $true "Service '$ServiceName' is installed and Running." ''
}

try {
    New-Item -ItemType Directory -Path $OutputDirectory -Force -ErrorAction Stop | Out-Null
    Write-PrerequisiteLog "Prerequisite gate started. Computer=$computerName Context=$env:USERDOMAIN\$env:USERNAME"
    $results = @(
        Get-EntraJoinResult
        Get-IntuneEnrollmentResult
        Get-RunningServiceResult 'Intune Management Extension' 'IntuneManagementExtension' 'Install or repair Microsoft Intune Management Extension and ensure the service is running.'
        Get-RunningServiceResult 'FSLogix' 'frxsvc' 'Install or repair FSLogix and ensure the frxsvc service is running.'
    )
    foreach ($result in $results) {
        Write-PrerequisiteLog "$($result.Prerequisite): $($result.Result) - $($result.Evidence)"
    }

    $failed = @($results | Where-Object { $_.Result -ne 'Pass' })
    $head = '<meta charset="utf-8"><style>body{font:14px "Segoe UI",sans-serif;margin:24px;color:#18212b}table{width:100%;border-collapse:collapse}th,td{padding:8px;border:1px solid #aeb9c4;text-align:left;vertical-align:top}th{background:#e6edf3}.pass{background:#d9f2df}.fail{background:#f8d7da}</style>'
    $rows = foreach ($result in $results) {
        $class = if ($result.Result -eq 'Pass') { 'pass' } else { 'fail' }
        "<tr class='$class'><td>$(ConvertTo-HtmlText $result.Prerequisite)</td><td>$($result.Result)</td><td>$(ConvertTo-HtmlText $result.Evidence)</td><td>$(ConvertTo-HtmlText $result.Remediation)</td></tr>"
    }
    $body = "<h1>AVD UAT Prerequisite Report</h1><p>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Computer: $computerName | Context: $env:USERDOMAIN\$env:USERNAME</p><p>All four prerequisites must pass before UAT can run.</p><table><thead><tr><th>Prerequisite</th><th>Result</th><th>Evidence</th><th>Remediation</th></tr></thead><tbody>$($rows -join '')</tbody></table>"
    ConvertTo-Html -Title 'AVD UAT Prerequisite Report' -Head $head -Body $body | Set-Content -LiteralPath $reportPath -Encoding UTF8

    $gate = @{
        Passed = $failed.Count -eq 0
        Results = $results
        LogPath = $logPath
        ReportPath = $reportPath
    }
    Write-PrerequisiteLog "Prerequisite gate completed. Passed=$($gate.Passed)"
    if ($Integrated) { return $gate }
    Write-Output "Prerequisite log: $logPath"
    Write-Output "Prerequisite report: $reportPath"
    if ($gate.Passed) { exit 0 }
    exit 2
}
catch {
    try { Write-PrerequisiteLog "Prerequisite evaluation error: $($_.Exception.Message)" 'ERROR' } catch { }
    if ($Integrated) { throw }
    Write-Warning "Prerequisite evaluation error: $($_.Exception.Message)"
    exit 3
}

# SIG # Begin signature block
# MIIHMAYJKoZIhvcNAQcCoIIHITCCBx0CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDpEyPTn5O1+WZ/
# rSHq20M3b6ynl4EVK758JIuIR3yYW6CCBBwwggQYMIICgKADAgECAhARPahbg4yR
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
# MC8GCSqGSIb3DQEJBDEiBCBHq2wUl0AEdxddfToOC9u8tSxhoEpuvVMDQReDI5Nu
# rDANBgkqhkiG9w0BAQEFAASCAYBAEGwP9qTeBlah91zdQFfwmxum4ODGFvtT2qKs
# 0redJ/ihxSEeYLHKIYf219bK/nDCD61/OnInmuGQGC5mnkGx+AlaDyD/zVEli5Vh
# ZSGs15Mxi5u6dyx1s2iFgXO7OYDTiNNm5LoXyFFTGt6Iax90fe+u/iwOtls7cUh0
# Ok0L08pWVbEKum9iHFF0qZAdes3oCHzlckLA4DzzNvJrMcdObYbQ/cSwzw/K30nd
# AU8A8H3wB19wOWcCE1diBLFnZWJoFDd6022kbbZzn8uDQh7yKfVxh9yFlUIt5sx2
# bN4/z3zjRmUStlBt2R8vsTDpeA4mvv/n/vlYelisMB9cWtNLgDkdi6SNVI3jjxHu
# 7dGBsFEhZeJgWt3slD9z2zI+STR9uJQW2BrvXqFVwS+5Z9lH65mWd1dEMo6FY8ak
# FKXoR/1L6VooktItpcCRCnKTfbaiI4RGAcDouSa4xj8cMknzY05lxuzS7EVzNDUu
# wUkNvkNyYZuYVV7aa+KRhUch1I0=
# SIG # End signature block
