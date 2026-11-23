[CmdletBinding()]
param(
    [string]$AutomatedResultsCsv,
    [string]$InputFile,
    [string]$OutputDirectory,
    [string]$UatDirectory = 'C:\Program Files\ArmyAVD\UAT',
    [string]$WorkbookTool,
    [switch]$SkipUnresolved,
    [switch]$NoAssistedLaunch,
    [string]$CheckpointPath
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
$sourceTimestamp = (Get-Item -LiteralPath $InputFile -ErrorAction Stop).BaseName -replace '^AVD Checklist - Completed - ', ''
$pairedCsv = Join-Path (Split-Path -Parent $InputFile) "AVD Checklist Results - $sourceTimestamp.csv"
if ($AutomatedResultsCsv -and $AutomatedResultsCsv -ne $pairedCsv) { throw "Manual evidence must match the latest automated workbook: $pairedCsv" }
$AutomatedResultsCsv = $pairedCsv
if (-not $AutomatedResultsCsv -or -not (Test-Path -LiteralPath $AutomatedResultsCsv -PathType Leaf)) { throw 'Automated results CSV was not found. Run the automated workflow first or supply -AutomatedResultsCsv.' }
if (-not (Test-Path -LiteralPath $InputFile -PathType Leaf)) { throw "Input checklist not found: $InputFile" }
if (-not (Test-Path -LiteralPath $WorkbookTool -PathType Leaf)) { throw "Workbook helper not found: $WorkbookTool" }

. (Join-Path $executionRoot 'AVD-UAT-ManualActions.ps1')
$results = @(Import-Csv -LiteralPath $AutomatedResultsCsv)
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
if (-not $SkipUnresolved) {
    if (-not $CheckpointPath) {
        $checkpointDirectory = Join-Path $env:LOCALAPPDATA 'ArmyAVD\UAT\Manual'
        $CheckpointPath = Join-Path $checkpointDirectory "AVD Manual Progress - $sourceTimestamp.json"
    }
    if ($CheckpointPath -notmatch '\.json$') { throw 'CheckpointPath must name a .json progress file, not a checklist or evidence file.' }
    $checkpointParent = Split-Path -Parent $CheckpointPath
    if (-not $checkpointParent) { $checkpointParent = '.' }
    New-Item -ItemType Directory -Path $checkpointParent -Force | Out-Null
    $CheckpointPath = Join-Path (Convert-Path -LiteralPath $checkpointParent -ErrorAction Stop) (Split-Path -Leaf $CheckpointPath)
    Write-Host "Progress checkpoint: $CheckpointPath"
    try {
        $state = if (Test-Path -LiteralPath $CheckpointPath -PathType Leaf) {
            Read-ManualCheckpoint $CheckpointPath $InputFile $AutomatedResultsCsv $results
        } else {
            New-ManualCheckpoint $InputFile $AutomatedResultsCsv $results
        }
        Save-ManualCheckpoint $state $CheckpointPath
        Invoke-ManualReview $state $CheckpointPath -NoAssistedLaunch:$NoAssistedLaunch
        $results = @($state.Rows)
        $timestamp = $state.ReportTimestamp
    }
    catch {
        Write-Warning "Manual review did not finish: $($_.Exception.Message)"
        Write-Warning "Saved progress is retained at: $CheckpointPath. Rerun with the same -CheckpointPath to resume. No application was closed by this tool."
        throw
    }
}
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
if (-not $SkipUnresolved) {
    $state.Outputs = @($resultsCsv, $outputWorkbook, $outputHtml)
    Save-ManualCheckpoint $state $CheckpointPath
}

Write-Output 'AVD UAT manual review completed.'
Write-Output "HTML report:    $outputHtml"
Write-Output "Excel checklist: $outputWorkbook"
Write-Output "CSV evidence:   $resultsCsv"

