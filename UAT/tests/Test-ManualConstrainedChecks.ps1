[CmdletBinding()]
param(
    [string]$SourceDirectory,
    [string]$FixtureDirectory,
    [ValidateSet('Harness', 'Script', 'Helper')][string]$Mode = 'Harness',
    [switch]$SkipCompiled
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
if (-not $SourceDirectory) { $SourceDirectory = Split-Path -Parent $PSScriptRoot }

if ($Mode -eq 'Harness') {
    $originalUserProfile = $env:USERPROFILE
    $originalLocalAppData = $env:LOCALAPPDATA
    $root = Join-Path $env:TEMP ('AVD-Manual-CLM-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $root | Out-Null
    try {
        $deployment = Join-Path $root 'deployment'
        $automated = Join-Path $deployment 'Automated_Results_20261116-130000-000'
        New-Item -ItemType Directory -Path $automated -Force | Out-Null
        foreach ($name in @('AVD-UAT-Checklist-Manual.ps1', 'AVD-UAT-ManualActions.ps1', 'AVD-UAT-Paths.ps1')) {
            Copy-Item -LiteralPath (Join-Path $SourceDirectory $name) -Destination $deployment
        }
        $fileToolDirectory = Join-Path $deployment 'tools\AvdManualFileTool'
        New-Item -ItemType Directory -Path $fileToolDirectory -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $SourceDirectory 'tools\AvdManualFileTool\AvdManualFileTool.exe') -Destination $fileToolDirectory
        $workbookTool = Join-Path $SourceDirectory 'tools\AvdChecklistTool\AvdChecklistTool.exe'
        $book = Join-Path $automated 'AVD Checklist - Completed - 20261116-130000-000.xlsx'
        $csv = Join-Path $automated 'AVD Checklist Results - 20261116-130000-000.csv'
        Copy-Item -LiteralPath (Join-Path $SourceDirectory 'AVD Test Plan Template - DO NOT DELETE OR MOVE.xlsx') -Destination $book
        $export = Join-Path $root 'template.csv'
        & $workbookTool export --input $book --output $export
        if ($LASTEXITCODE -ne 0) { throw 'Fixture workbook export failed.' }
        Import-Csv -LiteralPath $export | Select-Object Sheet, ID, Category, Test,
            @{ Name = 'Expected'; Expression = { $_.'Expected Results' } },
            @{ Name = 'Result'; Expression = { 'Pass' } },
            @{ Name = 'IssueReference'; Expression = { '' } },
            @{ Name = 'Notes'; Expression = { 'Original automated evidence' } },
            @{ Name = 'AutomationReview'; Expression = { 'Original automated review' } } |
            Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
        $bookHash = (Get-FileHash -LiteralPath $book).Hash
        $csvHash = (Get-FileHash -LiteralPath $csv).Hash
        $checker = @'
if ($ExecutionContext.SessionState.LanguageMode -ne 'ConstrainedLanguage') { throw 'Launcher is not executing in CLM.' }
@{ Passed = $true; LogPath = 'fixture.log'; ReportPath = 'fixture.html' }
'@
        Set-Content -LiteralPath (Join-Path $deployment 'AVD-UAT-Prerequisite-Check.ps1') -Value $checker -Encoding UTF8
        $mocks = @'
function Get-FileHash { throw 'Get-FileHash must not be required.' }
function Read-Host {
    param([string]$Prompt)
    if ($ExecutionContext.SessionState.LanguageMode -ne 'ConstrainedLanguage') { throw 'Manual prompt escaped CLM.' }
    if ($Prompt -like 'Final result:*') { return 'P' }
    if ($Prompt -like 'Manual evidence*') { return 'Observed in CLM fixture' }
    if ($Prompt -like '*onfirm test windows*') { return 'C' }
    throw "Unexpected prompt: $Prompt"
}
function Start-Process {
    param([string]$FilePath, [string[]]$ArgumentList, [string]$ErrorAction)
    if ($ExecutionContext.SessionState.LanguageMode -ne 'ConstrainedLanguage') { throw 'Assisted launch escaped CLM.' }
    Add-Content -LiteralPath (Join-Path $env:LOCALAPPDATA 'launches.log') -Value $FilePath
}
function Test-Path {
    param([string]$LiteralPath, [string]$PathType)
    if ($LiteralPath -match '\.(exe|lnk|msc)$' -and $LiteralPath -notlike '*Avd*Tool.exe') { return $true }
    Microsoft.PowerShell.Management\Test-Path @PSBoundParameters
}
function Get-AppxPackage {
    param([string]$Name, [string]$ErrorAction)
    $programFiles = if ($env:ProgramW6432) { $env:ProgramW6432 } else { $env:ProgramFiles }
    @{ Version = '11.10.0.0'; InstallLocation = "$programFiles\WindowsApps\$($Name)_11.10.0.0_x64__fixture" }
}
'@
        # Test doubles stay in the same constrained script as the production helper.
        Add-Content -LiteralPath (Join-Path $deployment 'AVD-UAT-ManualActions.ps1') -Value $mocks -Encoding UTF8
        foreach ($runtime in @('powershell.exe', 'pwsh.exe')) {
            foreach ($mode in @('Helper', 'Script')) {
                $case = Join-Path $root "$runtime-$mode"
                New-Item -ItemType Directory -Path $case | Out-Null
                & $runtime -NoProfile -NonInteractive -File $PSCommandPath -SourceDirectory $deployment -FixtureDirectory $case -Mode $mode
                if ($LASTEXITCODE -ne 0) { throw "$runtime $mode CLM test failed with exit code $LASTEXITCODE." }
            }
        }
        if (-not $SkipCompiled) {
            $compiledSource = Join-Path $root 'compiled-clm.ps1'
            $source = (Get-Content -LiteralPath (Join-Path $deployment 'AVD-UAT-Checklist-Manual.ps1') -Raw) -replace '(?ms)^# SIG # Begin signature block.*\z', ''
            $source = $source.Replace('Set-StrictMode -Version 2.0', '$ExecutionContext.SessionState.LanguageMode = ''ConstrainedLanguage''; Set-StrictMode -Version 2.0')
            Set-Content -LiteralPath $compiledSource -Value $source -Encoding UTF8
            $exe = Join-Path $deployment 'compiled-clm.exe'
            Invoke-ps2exe -inputFile $compiledSource -outputFile $exe -x64 -STA -ErrorAction Stop | Out-Null
            $env:USERPROFILE = Join-Path $root 'compiled-profile'
            $env:LOCALAPPDATA = Join-Path $root 'compiled-local'
            New-Item -ItemType Directory -Path $env:LOCALAPPDATA | Out-Null
            $checkpoint = Join-Path $env:LOCALAPPDATA 'compiled.json'
            & $exe -UatDirectory $deployment -WorkbookTool $workbookTool -CheckpointPath $checkpoint | Out-Null
            if ($LASTEXITCODE -ne 0) { throw 'Compiled CLM workflow failed.' }
            $saved = Get-Content -LiteralPath $checkpoint -Raw | ConvertFrom-Json
            if ($saved.Outputs.Count -ne 3 -or $saved.WorkbookHash -cne $bookHash -or $saved.CsvHash -cne $csvHash) { throw 'Compiled CLM checkpoint/evidence is incomplete.' }
            $launches = @(Get-Content -LiteralPath (Join-Path $env:LOCALAPPDATA 'launches.log')).Count
            & $exe -UatDirectory $deployment -WorkbookTool $workbookTool -CheckpointPath $checkpoint | Out-Null
            if ($LASTEXITCODE -ne 0) { throw 'Compiled CLM resume failed.' }
            if (@(Get-Content -LiteralPath (Join-Path $env:LOCALAPPDATA 'launches.log')).Count -ne $launches) { throw 'Compiled CLM resume reopened applications.' }
            Write-Output 'PASS: compiled CLM review, SHA-256, atomic checkpoints, three reports, and resume.'
        }
        if ((Get-FileHash -LiteralPath $book).Hash -cne $bookHash -or (Get-FileHash -LiteralPath $csv).Hash -cne $csvHash) { throw 'Automated source evidence was modified.' }
        Write-Output 'PASS: manual CLM scenarios; host services, registry, and application launches are simulated.'
    }
    finally {
        $env:USERPROFILE = $originalUserProfile
        $env:LOCALAPPDATA = $originalLocalAppData
        Remove-Item -LiteralPath $root -Recurse -Force
    }
    return
}

$env:USERPROFILE = Join-Path $FixtureDirectory 'profile'
$env:LOCALAPPDATA = Join-Path $FixtureDirectory 'local'
$env:TEMP = Join-Path $FixtureDirectory 'temp'
New-Item -ItemType Directory -Path $env:LOCALAPPDATA, $env:TEMP -Force | Out-Null
# The real workbook utility is supplied independently of the fixture installation.
$workbookTool = Join-Path (Split-Path -Parent $PSScriptRoot) 'tools\AvdChecklistTool\AvdChecklistTool.exe'
$automated = Join-Path $SourceDirectory 'Automated_Results_20261116-130000-000'
$book = Join-Path $automated 'AVD Checklist - Completed - 20261116-130000-000.xlsx'
$csv = Join-Path $automated 'AVD Checklist Results - 20261116-130000-000.csv'
$checkpoint = Join-Path $env:LOCALAPPDATA 'progress.json'

$ExecutionContext.SessionState.LanguageMode = 'ConstrainedLanguage'
if ($Mode -eq 'Script') {
    & (Join-Path $SourceDirectory 'AVD-UAT-Checklist-Manual.ps1') -UatDirectory $SourceDirectory -WorkbookTool $workbookTool -CheckpointPath $checkpoint | Out-Null
    $saved = Get-Content -LiteralPath $checkpoint -Raw | ConvertFrom-Json
    if ($saved.Outputs.Count -ne 3 -or $saved.ClosurePending) { throw 'CLM final checkpoint is incomplete.' }
    foreach ($output in $saved.Outputs) { if (-not (Test-Path -LiteralPath $output)) { throw "Missing CLM output: $output" } }
    $rows = @(Import-Csv -LiteralPath $saved.Outputs[0])
    if ($rows.Count -ne 99 -or @($rows | Where-Object Notes -like '*Observed in CLM fixture*').Count -lt 60) { throw 'CLM reports lost manual observations.' }
    $opens = @(Get-Content -LiteralPath (Join-Path $env:LOCALAPPDATA 'launches.log'))
    foreach ($suffix in @('Microsoft.Paint_11.10.0.0_x64__fixture\PaintApp\mspaint.exe', 'Microsoft.ScreenSketch_11.10.0.0_x64__fixture\SnippingTool\SnippingTool.exe', 'Microsoft.Windows.Photos_11.10.0.0_x64__fixture\Photos.exe')) {
        if (@($opens | Where-Object { $_.EndsWith($suffix) }).Count -ne 1) { throw "Packaged app was not launched once in CLM: $suffix" }
    }
    & (Join-Path $SourceDirectory 'AVD-UAT-Checklist-Manual.ps1') -UatDirectory $SourceDirectory -WorkbookTool $workbookTool -CheckpointPath $checkpoint | Out-Null
    if (@(Get-Content -LiteralPath (Join-Path $env:LOCALAPPDATA 'launches.log')).Count -ne $opens.Count) { throw 'CLM resume reopened applications.' }
    & (Join-Path $SourceDirectory 'AVD-UAT-Checklist-Manual.ps1') -UatDirectory $SourceDirectory -WorkbookTool $workbookTool -SkipUnresolved | Out-Null
    Write-Output 'PASS: CLM launcher, packaged apps, results, three reports, completed resume, and skip.'
    return
}

. (Join-Path $SourceDirectory 'AVD-UAT-ManualActions.ps1')
$rows = @(Import-Csv -LiteralPath $csv | Where-Object { $_.Sheet -eq 'Azure Virtual Desktop Tasks' -and $_.ID -in @('AVD-54', 'AVD-55') })
$state = New-ManualCheckpoint $book $csv $rows
Save-ManualCheckpoint $state $checkpoint
$originalHash = $state.WorkbookHash
$state.ClosurePending = $true
$state.ActiveGroup = 'Azure Virtual Desktop Tasks|Snipping Tool'
$state.ActiveApplication = 'Snipping Tool'
Save-ManualCheckpoint $state $checkpoint
$restored = Read-ManualCheckpoint $checkpoint $book $csv $rows
if ($restored.ClosurePending -or $restored.WorkbookHash -cne $originalHash) { throw 'Interrupted CLM checkpoint resume failed.' }
Invoke-ManualReview $restored $checkpoint
$saved = Get-Content -LiteralPath $checkpoint -Raw | ConvertFrom-Json
if ($saved.Steps.Count -ne 2 -or $saved.CompletedRows.Count -ne 2 -or $saved.ClosurePending) { throw 'CLM grouped review did not persist observations.' }
$saved.CsvHash = 'incorrect'
Save-ManualCheckpoint $saved $checkpoint
$rejected = $false
try { Read-ManualCheckpoint $checkpoint $book $csv $rows | Out-Null }
catch {
    if ($_.Exception.Message -notlike '*CsvHash does not match*') { throw }
    $rejected = $true
}
if (-not $rejected) { throw 'CLM accepted mismatched evidence.' }
if (@(Get-ChildItem -LiteralPath $env:LOCALAPPDATA -Filter '*.tmp').Count -ne 0) { throw 'CLM checkpoint write leaked temporary files.' }
Write-Output 'PASS: CLM grouping, checkpoint replacement, interrupted resume, saved observations, and evidence mismatch rejection.'
