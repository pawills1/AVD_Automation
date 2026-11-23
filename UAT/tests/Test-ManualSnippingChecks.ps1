[CmdletBinding()]
param([string]$EnginePath)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
if (-not $EnginePath) { $EnginePath = Join-Path (Split-Path -Parent $PSScriptRoot) 'AVD-UAT-Checklist.ps1' }
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($EnginePath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw "Engine parse errors: $parseErrors" }
$definitions = $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false)
foreach ($name in @('New-Result', 'Get-ValidationCheckDescription', 'Get-AutomatedResult')) {
    $definition = @($definitions | Where-Object Name -eq $name)
    if ($definition.Count -ne 1) { throw "Expected one function named $name." }
    . ([scriptblock]::Create($definition[0].Extent.Text))
}

$script:probeCalls = 0
function Stop-UnexpectedProbe {
    $script:probeCalls++
    throw 'Manual-only Snipping Tool rows must not inspect user aliases or launch applications.'
}
function Get-SnippingToolResult { Stop-UnexpectedProbe }
function Get-SnippingToolPath { Stop-UnexpectedProbe }
function Get-ApplicationStartupResult { Stop-UnexpectedProbe }
function Invoke-BoundedApplicationCheck { Stop-UnexpectedProbe }
function Get-ChildItem { Stop-UnexpectedProbe }
function Test-Path { Stop-UnexpectedProbe }
function Start-Process { Stop-UnexpectedProbe }

$passed = 0
foreach ($isSystem in @($true, $false)) {
    foreach ($RunInteractiveChecks in @($false, $true)) {
        foreach ($id in @('AVD-54', 'AVD-55')) {
            $result = Get-AutomatedResult -Sheet 'Azure Virtual Desktop Tasks' -Id $id
            if ($result.Result -ne 'Manual') { throw "$id returned '$($result.Result)' instead of Manual." }
            if ($result.Notes -notlike '*signed-in UAT user*' -or $result.Review -notlike 'Manual task:*') {
                throw "$id is missing actionable signed-in-user guidance."
            }
            $description = Get-ValidationCheckDescription 'Azure Virtual Desktop Tasks' $id $result.Review
            if ($description -notlike 'Manual only:*' -or $description -like '*WindowsApps*') {
                throw "$id still describes an automated application probe."
            }
            $passed++
        }
        foreach ($id in @('AVD-05', 'AVD-21', 'AVD-25', 'AVD-41', 'AVD-58', 'AVD-59', 'AVD-61', 'AVD-62', 'AVD-69', 'AVD-80')) {
            if ((Get-AutomatedResult -Sheet 'Azure Virtual Desktop Tasks' -Id $id).Result -ne 'Manual') {
                throw "Existing manual-only row $id changed behavior."
            }
            $passed++
        }
    }
}
foreach ($id in @('AVD-54', 'AVD-55')) {
    if ($null -ne (Get-AutomatedResult -Sheet 'Remote Apps Tasks' -Id $id)) { throw "Unexpected Remote Apps mapping for $id." }
}
if ($script:probeCalls -ne 0) { throw "Attempted $script:probeCalls forbidden application probes." }
Write-Output "PASS: $passed manual-only routing scenarios; zero user-alias searches or application probes; worksheet scoping preserved."
