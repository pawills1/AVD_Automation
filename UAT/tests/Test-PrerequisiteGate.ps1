[CmdletBinding()]
param(
    [string]$SourceDirectory,
    [ValidateSet('Harness', 'Integrated', 'Standalone', 'Automated', 'Manual')]
    [string]$Mode = 'Harness',
    [string]$OutputDirectory,
    [string]$ActualPool = 'P-ARMY-AVD-AZ-INT01B-1',
    [switch]$Constrained
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
if (-not $SourceDirectory) { $SourceDirectory = Split-Path -Parent $PSScriptRoot }

if ($Mode -eq 'Harness') {
    $testRoot = Join-Path $env:TEMP ('AVD-HostPool-GateTests-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $testRoot | Out-Null
    $passed = 0
    try {
        foreach ($runtime in @('powershell.exe', 'pwsh.exe')) {
            foreach ($language in @('FullLanguage', 'ConstrainedLanguage')) {
                foreach ($case in @(
                    @('Integrated', '1', 0),
                    @('Integrated', '2', 0),
                    @('Standalone', '1', 0),
                    @('Standalone', '2', 2),
                    @('Automated', '2', 2),
                    @('Manual', '2', 2)
                )) {
                    $caseName = "$runtime-$language-$($case[0])-$($case[1])"
                    $caseDirectory = Join-Path $testRoot $caseName
                    New-Item -ItemType Directory -Path $caseDirectory | Out-Null
                    $arguments = @(
                        '-NoProfile', '-NonInteractive', '-File', $PSCommandPath,
                        '-SourceDirectory', $SourceDirectory, '-Mode', $case[0],
                        '-OutputDirectory', $caseDirectory,
                        '-ActualPool', "P-ARMY-AVD-AZ-INT01B-$($case[1])"
                    )
                    if ($language -eq 'ConstrainedLanguage') { $arguments += '-Constrained' }
                    $ErrorActionPreference = 'Continue'
                    $output = & $runtime @arguments 2>&1
                    $exitCode = $LASTEXITCODE
                    $ErrorActionPreference = 'Stop'
                    if ($exitCode -ne $case[2]) { throw "$caseName exited $exitCode instead of $($case[2]): $output" }
                    $reports = @(Get-ChildItem -LiteralPath $caseDirectory -Recurse -Filter 'AVD UAT Prerequisite Report*.html')
                    $logs = @(Get-ChildItem -LiteralPath $caseDirectory -Recurse -Filter 'AVD UAT Prerequisite Log*.log')
                    if ($reports.Count -ne 1 -or $logs.Count -ne 1) { throw "$caseName did not produce one prerequisite report and log." }
                    $html = Get-Content -LiteralPath $reports[0].FullName -Raw
                    $log = Get-Content -LiteralPath $logs[0].FullName -Raw
                    foreach ($text in @('All 6 prerequisites', 'Host Pool Status', 'Device=AVDAZINT01B1001', 'Expected pool=P-ARMY-AVD-AZ-INT01B-1')) {
                        if ($html -notlike "*$text*") { throw "$caseName report is missing '$text'." }
                    }
                    $expectedPassed = $case[1] -eq '1'
                    if ($log -notlike "*Passed=$expectedPassed*") { throw "$caseName has an incorrect gate result: $log" }
                    if (-not $expectedPassed -and ($html -notlike '*host-pool number mismatch*' -or $log -notlike '*Host Pool Status: Fail*')) {
                        throw "$caseName did not report the pool mismatch."
                    }
                    if ($case[0] -eq 'Integrated' -and ($output -join "`n") -notlike '*Integrated gate verified*') { throw "$caseName did not verify integrated results." }
                    $passed++
                    Write-Output "PASS: $caseName (exit $exitCode)"
                }
            }
        }
        Write-Output "PASS: $passed gate/report/launcher scenarios; no live registry or machine configuration was changed."
    }
    finally {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
    return
}

# Simulate host dependencies only inside the child process; execute the actual gate and launchers.
$env:COMPUTERNAME = 'AVDAZINT01B1001'
$env:TEMP = $OutputDirectory
$global:mockHostPoolValue = $ActualPool
function Test-Path {
    param([string]$LiteralPath, [string]$PathType)
    if ($LiteralPath -match '\\(?:System32|Sysnative)\\(?:dsregcmd|schtasks)\.exe\z') { return $false }
    Microsoft.PowerShell.Management\Test-Path @PSBoundParameters
}
function dsregcmd.exe {
    $global:LASTEXITCODE = 0
    'AzureAdJoined : YES'
}
function schtasks.exe {
    $global:LASTEXITCODE = 0
    '\Microsoft\Windows\EnterpriseMgmt\11111111-1111-1111-1111-111111111111\Schedule #1'
}
function Get-ChildItem {
    param([string]$Path, [string]$ErrorAction)
    if ($Path -ne 'HKLM:\SOFTWARE\Microsoft\Enrollments') { throw "Unexpected inventory access after the gate: $Path" }
    @{ PSChildName = '11111111-1111-1111-1111-111111111111'; PSPath = 'mock-enrollment' }
}
function Get-ItemPropertyValue {
    param([string]$LiteralPath, [string]$Name, [string]$ErrorAction)
    if ($LiteralPath -eq 'mock-enrollment' -and $Name -eq 'ProviderID') { return 'MS DM Server' }
    if ($LiteralPath -eq 'HKLM:\SOFTWARE\Microsoft\RDMonitoringAgent' -and $Name -eq 'SessionHostPool') { return $global:mockHostPoolValue }
    throw "Unexpected registry access: $LiteralPath\$Name"
}
function Get-Service {
    param([string]$Name, [string]$ErrorAction)
    if ($Name -notin @('IntuneManagementExtension', 'frxsvc', 'MicrosoftDependencyAgent')) { throw "Unexpected service: $Name" }
    @{ Status = 'Running' }
}
if ($Constrained) { $ExecutionContext.SessionState.LanguageMode = 'ConstrainedLanguage' }
$checker = Join-Path $SourceDirectory 'AVD-UAT-Prerequisite-Check.ps1'
switch ($Mode) {
    'Integrated' {
        $gate = & $checker -OutputDirectory $OutputDirectory -Integrated
        if ($gate.Results.Count -ne 6) { throw 'Expected six prerequisite results.' }
        if ($gate.Passed -ne ($ActualPool -eq 'P-ARMY-AVD-AZ-INT01B-1')) {
            throw "Incorrect integrated gate result: $($gate.Results | ForEach-Object { $_.Evidence })"
        }
        Write-Output 'Integrated gate verified'
    }
    'Standalone' {
        & $checker -OutputDirectory $OutputDirectory
        exit $LASTEXITCODE
    }
    'Automated' {
        & (Join-Path $SourceDirectory 'AVD-UAT-Checklist-Automated.ps1') -OutputDirectory $OutputDirectory
        exit $LASTEXITCODE
    }
    'Manual' {
        & (Join-Path $SourceDirectory 'AVD-UAT-Checklist-Manual.ps1')
        exit $LASTEXITCODE
    }
}
