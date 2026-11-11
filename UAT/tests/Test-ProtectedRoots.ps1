[CmdletBinding()]
param([string]$EnginePath, [switch]$Constrained)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
if (-not $EnginePath) { $EnginePath = Join-Path (Split-Path -Parent $PSScriptRoot) 'AVD-UAT-Checklist.ps1' }
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($EnginePath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw "Engine parse errors: $parseErrors" }
$definitions = $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false)
$functionText = @()
foreach ($name in @('New-Result', 'Get-ValidationCheckDescription', 'Get-ProtectedRootsResult', 'Get-AutomatedResult')) {
    $definition = @($definitions | Where-Object Name -eq $name)
    if ($definition.Count -ne 1) { throw "Expected one function named $name." }
    $functionText += $definition[0].Extent.Text
}
if ($Constrained) { $ExecutionContext.SessionState.LanguageMode = 'ConstrainedLanguage' }
foreach ($text in $functionText) { Invoke-Expression $text }

$script:readCalls = 0
function Get-ItemPropertyValue {
    param([string]$LiteralPath, [string]$Name, [string]$ErrorAction)
    if ($LiteralPath -cne 'HKLM:\Software\Policies\Microsoft\SystemCertificates\Root\ProtectedRoots' -or $ErrorAction -ne 'Stop') {
        throw "Unexpected registry access: $LiteralPath ($ErrorAction)"
    }
    $script:readCalls++
    if ($script:case.Error -eq 'Key' -or $script:case.Error -eq $Name) { throw "Simulated missing or unreadable $($script:case.Error)" }
    switch ($Name) {
        'Flags' { return $script:case.Flags }
        'PeerUsages' { return ,$script:case.Usages }
        default { throw "Unexpected registry value: $Name" }
    }
}
function Get-ChildItem { throw 'AVD-81 must not inspect certificate stores.' }
function Set-ItemProperty { throw 'AVD-81 must not change registry policy.' }

$expected = [string[]]@('1.3.6.1.5.5.7.3.2', '1.3.6.1.5.5.7.3.4', '1.3.6.1.4.1.311.10.3.4')
$cases = @(
    @{ Name = 'Exact'; Flags = 65537; Usages = $expected; Error = ''; Result = 'Pass' }
    @{ Name = 'Reordered'; Flags = 65537; Usages = [string[]]@($expected[2], $expected[0], $expected[1]); Error = ''; Result = 'Pass' }
    @{ Name = 'Wrong flags'; Flags = 1; Usages = $expected; Error = ''; Result = 'Fail' }
    @{ Name = 'String flags'; Flags = '65537'; Usages = $expected; Error = ''; Result = 'Fail' }
    @{ Name = 'QWORD flags'; Flags = [long]65537; Usages = $expected; Error = ''; Result = 'Fail' }
    @{ Name = 'String usages'; Flags = 65537; Usages = ($expected -join ', '); Error = ''; Result = 'Fail' }
    @{ Name = 'Missing usage'; Flags = 65537; Usages = [string[]]@($expected[0], $expected[1]); Error = ''; Result = 'Fail' }
    @{ Name = 'Extra usage'; Flags = 65537; Usages = [string[]]@($expected + '1.2.3'); Error = ''; Result = 'Fail' }
    @{ Name = 'Duplicate usage'; Flags = 65537; Usages = [string[]]@($expected[0], $expected[1], $expected[1]); Error = ''; Result = 'Fail' }
    @{ Name = 'Wrong usage'; Flags = 65537; Usages = [string[]]@($expected[0], $expected[1], '1.2.3'); Error = ''; Result = 'Fail' }
    @{ Name = 'Empty usages'; Flags = 65537; Usages = [string[]]@(); Error = ''; Result = 'Fail' }
    @{ Name = 'Missing key'; Flags = 65537; Usages = $expected; Error = 'Key'; Result = 'Fail' }
    @{ Name = 'Missing flags'; Flags = 65537; Usages = $expected; Error = 'Flags'; Result = 'Fail' }
    @{ Name = 'Unreadable usages'; Flags = 65537; Usages = $expected; Error = 'PeerUsages'; Result = 'Fail' }
)
$passed = 0
foreach ($isSystem in @($true, $false)) {
    foreach ($RunInteractiveChecks in @($false, $true)) {
        foreach ($script:case in $cases) {
            $result = Get-AutomatedResult -Sheet 'Azure Virtual Desktop Tasks' -Id 'AVD-81'
            if ($result.Result -ne $script:case.Result) { throw "$($script:case.Name): expected $($script:case.Result), got $($result.Result): $($result.Notes)" }
            if ($result.Notes -notlike '*ProtectedRoots*') { throw 'Missing registry-path evidence.' }
            if ($script:case.Error -and $result.Notes -notlike "*Simulated missing or unreadable $($script:case.Error)*") { throw 'Missing read-error details.' }
            if ($result.Review -ne 'Automated') { throw 'AVD-81 unexpectedly requires manual review.' }
            $passed++
        }
    }
}
$callsBefore = $script:readCalls
if ($null -ne (Get-AutomatedResult -Sheet 'Remote Apps Tasks' -Id 'AVD-81')) { throw 'Unexpected Remote Apps mapping.' }
if ($script:readCalls -ne $callsBefore) { throw 'Remote Apps AVD-81 read registry policy.' }
$description = Get-ValidationCheckDescription 'Azure Virtual Desktop Tasks' 'AVD-81' 'Automated'
foreach ($text in @('ProtectedRoots', '65537', 'REG_DWORD', 'REG_MULTI_SZ') + $expected) {
    if (-not $description.Contains($text)) { throw "Missing validation description: $text" }
}
Write-Output "PASS: $passed Protected Roots scenarios ($($ExecutionContext.SessionState.LanguageMode)); worksheet scoping and report description verified. No machine configuration was changed."
