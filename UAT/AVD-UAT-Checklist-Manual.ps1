[CmdletBinding()]
param(
    [string]$AutomatedResultsCsv,
    [string]$InputFile,
    [string]$OutputDirectory,
    [string]$UatDirectory = 'C:\Program Files\ArmyAVD\UAT',
    [string]$WorkbookTool,
    [switch]$SkipUnresolved
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$executionRoot = if ($PSScriptRoot) { $PSScriptRoot } else { [AppDomain]::CurrentDomain.BaseDirectory.TrimEnd('\') }
. (Join-Path $executionRoot 'AVD-UAT-Paths.ps1')
$prerequisiteChecker = Join-Path $executionRoot 'AVD-UAT-Prerequisite-Check.ps1'
if (-not (Test-Path -LiteralPath $prerequisiteChecker -PathType Leaf)) {
    Write-Warning "Prerequisite gate error: checker not found: $prerequisiteChecker"
    exit 3
}
$prerequisiteOutput = Join-Path $env:TEMP 'ArmyAVD\UAT-Prerequisites'
try { $prerequisiteGate = & $prerequisiteChecker -OutputDirectory $prerequisiteOutput -Integrated }
catch {
    Write-Warning "Prerequisite gate error: $($_.Exception.Message)"
    exit 3
}
Write-Output "Prerequisite log: $($prerequisiteGate.LogPath)"
Write-Output "Prerequisite report: $($prerequisiteGate.ReportPath)"
if (-not $prerequisiteGate.Passed) {
    Write-Warning 'Prerequisite gate failed. Manual UAT was not started.'
    exit 2
}
if ($env:USERNAME -eq 'SYSTEM') { throw 'Manual checks must run as the signed-in UAT user, not SYSTEM.' }
$latestWorkbook = Resolve-UatAutomatedWorkbook -UatDirectory $UatDirectory
if ($InputFile -and $InputFile -ne $latestWorkbook) { throw "Manual input must be the latest automated workbook: $latestWorkbook" }
$InputFile = $latestWorkbook
$manualDirectory = Join-Path $env:USERPROFILE 'OneDrive - US Army\Documents'
if ($OutputDirectory -and $OutputDirectory.TrimEnd('\') -ne $manualDirectory.TrimEnd('\')) { throw "Manual final outputs must go to: $manualDirectory" }
$OutputDirectory = $manualDirectory
if (-not $WorkbookTool) { $WorkbookTool = Join-Path $executionRoot 'tools\AvdChecklistTool\AvdChecklistTool.exe' }
$sourceTimestamp = [IO.Path]::GetFileNameWithoutExtension($InputFile) -replace '^AVD Checklist - Completed - ', ''
$pairedCsv = Join-Path (Split-Path -Parent $InputFile) "AVD Checklist Results - $sourceTimestamp.csv"
if ($AutomatedResultsCsv -and $AutomatedResultsCsv -ne $pairedCsv) { throw "Manual evidence must match the latest automated workbook: $pairedCsv" }
$AutomatedResultsCsv = $pairedCsv
if (-not $AutomatedResultsCsv -or -not (Test-Path -LiteralPath $AutomatedResultsCsv -PathType Leaf)) { throw 'Automated results CSV was not found. Run the automated workflow first or supply -AutomatedResultsCsv.' }
if (-not (Test-Path -LiteralPath $InputFile -PathType Leaf)) { throw "Input checklist not found: $InputFile" }
if (-not (Test-Path -LiteralPath $WorkbookTool -PathType Leaf)) { throw "Workbook helper not found: $WorkbookTool" }

function Add-ManualText {
    param([string]$Existing, [string]$Label, [string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $Existing }
    if ([string]::IsNullOrWhiteSpace($Existing)) { return "$Label$Value" }
    return "$Existing | $Label$Value"
}

$results = @(Import-Csv -LiteralPath $AutomatedResultsCsv)
$unresolved = @($results | Where-Object { $_.Result -in @('Partial Pass', 'Manual') })
foreach ($check in $unresolved) {
    if ($SkipUnresolved) { continue }
    Write-Host "`n$($check.Sheet) / $($check.ID): $($check.Test)" -ForegroundColor Cyan
    Write-Host "Expected: $($check.Expected)"
    Write-Host "Automated evidence: $($check.Notes)" -ForegroundColor DarkCyan
    do { $choice = (Read-Host 'Final result: [P]ass, [F]ail, [M]anual/unresolved, [S]kip').Trim().ToUpperInvariant() } until ($choice -in @('P', 'F', 'M', 'S'))
    if ($choice -eq 'S') { continue }
    $check.Result = switch ($choice) { 'P' { 'Pass' }; 'F' { 'Fail' }; default { 'Manual' } }
    $manualNotes = Read-Host 'Manual evidence or notes'
    $check.Notes = Add-ManualText $check.Notes 'Manual evidence: ' $manualNotes
    if ($choice -eq 'F') { $check.IssueReference = Read-Host 'Issue reference (optional)' }
    $check.AutomationReview = Add-ManualText $check.AutomationReview 'Manual review: ' "Completed by $env:USERDOMAIN\$env:USERNAME on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
}

New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
$resultsCsv = Join-Path $OutputDirectory "AVD Checklist Final Results - $timestamp.csv"
$outputWorkbook = Join-Path $OutputDirectory "AVD Checklist - Final - $timestamp.xlsx"
$outputHtml = Join-Path $OutputDirectory "AVD UAT Final Report - $timestamp.html"
$results | Export-Csv -LiteralPath $resultsCsv -NoTypeInformation -Encoding UTF8
& $WorkbookTool update --input $InputFile --results $resultsCsv --output $outputWorkbook
if ($LASTEXITCODE -ne 0) { throw "Workbook update failed with exit code $LASTEXITCODE." }

$columns = @('Sheet', 'ID', 'Category', 'Test', 'Expected', 'Result', 'IssueReference', 'Notes', 'AutomationReview')
$groups = foreach ($status in @('Pass', 'Partial Pass', 'Manual', 'Fail')) {
    $items = @($results | Where-Object Result -eq $status)
    $class = if ($status -eq 'Pass') { 'pass' } elseif ($status -eq 'Fail') { 'fail' } else { 'manual' }
    $table = ($items | Select-Object $columns | ConvertTo-Html -Fragment -PreContent "<h2>$status ($($items.Count))</h2>") -replace '<table>', "<table class=`"status-table $class`">"
    $table -join "`r`n"
}
$body = "<header><h1>AVD UAT Final Report</h1><p>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Computer: $env:COMPUTERNAME | Reviewer: $env:USERDOMAIN\$env:USERNAME</p></header><main>$($groups -join "`r`n")</main>"
$head = '<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><style>body{margin:0;background:#f4f6f8;color:#18212b;font:14px "Segoe UI",sans-serif}header{padding:24px 3%;background:#17324d;color:#fff}h1{margin:0 0 6px}main{padding:8px 3% 32px}h2{margin-top:28px}table{width:100%;border-collapse:collapse}th,td{padding:8px;border:1px solid #aeb9c4;text-align:left;vertical-align:top}th{background:#e6edf3}.status-table.pass td{background:#d9f2df}.status-table.manual td{background:#fff0c2}.status-table.fail td{background:#f8d7da}@media(max-width:800px){table{display:block;overflow-x:auto}}</style>'
ConvertTo-Html -Title 'AVD UAT Final Report' -Head $head -Body $body | Set-Content -LiteralPath $outputHtml -Encoding UTF8

Write-Output 'AVD UAT manual review completed.'
Write-Output "HTML report:    $outputHtml"
Write-Output "Excel checklist: $outputWorkbook"
Write-Output "CSV evidence:   $resultsCsv"

# SIG # Begin signature block
# MIIHMAYJKoZIhvcNAQcCoIIHITCCBx0CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA6XIPBVslmyHrU
# N7BOWdX83pRgDx6g3BKeYb3ShxDLRqCCBBwwggQYMIICgKADAgECAhARPahbg4yR
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
# MC8GCSqGSIb3DQEJBDEiBCBA7algW+v1msc3r8eDPSqccV9kDp5Xp7SO7/gqq45y
# XDANBgkqhkiG9w0BAQEFAASCAYCEWOiB3GIWgVcFhFknoBfp5wiUQJRMHbY7k/JN
# 0ZnY/wfLd8tvEd1xrKOP6SPpxnEbfeIhe9l4KiOijU6+WTWjMUN81ysfKXqjp7Fi
# CcHcZMxy4lQe3hVtOx+dX2HkOPZZWLmbaMqJbxxtd61z+lgwyqYXrXRYdFgWZTdV
# ZfmbfmYmYzTyWaUiipX+1kqa74fXBunf+WBCzK+kK8XXI6PLUdPATTd6j1seTNkI
# U5dop6EukHpzH0SVtUh8nTc1rj2KI9tmDM/lCy28Rbsblh3s83UkCPmpBBhUT6gX
# vIHXKTIrQeWGE8EUUMqsfCdyAjUZMCCK9sKGhtntnk4Lrwj4o03NWPZoG4TzuRh7
# gu06G3Se4WK2IZm5qMi38nT7sE2DW/JntvQtwRFeKSKNVo3sFrmiOqyJMMscjGtO
# QgHtDp7tvZcLSdQLy46EOaFxbyKBxI+Qb6y6Hh4OWRz+7bQRF3Jd2SGNYTjm/c2f
# 0DX5N4dnwK9UPnJFuyO+uP7oR9Y=
# SIG # End signature block
