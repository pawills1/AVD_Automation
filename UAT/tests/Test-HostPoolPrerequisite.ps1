[CmdletBinding()]
param(
    [string]$CheckerPath
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
if (-not $CheckerPath) { $CheckerPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'AVD-UAT-Prerequisite-Check.ps1' }
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($CheckerPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw "Checker parse errors: $parseErrors" }
$definitions = $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false)
foreach ($definition in $definitions) {
    . ([scriptblock]::Create($definition.Extent.Text))
}
if (-not (Get-Command Get-HostPoolResult -ErrorAction SilentlyContinue)) { throw 'Host Pool Status prerequisite is not implemented.' }

$script:assertions = 0
function Assert-Equal {
    param($Actual, $Expected, [string]$Description)
    if ($Actual -cne $Expected) { throw "$Description`: expected '$Expected', got '$Actual'." }
    $script:assertions++
}

$script:actualPool = ''
$script:registryError = ''
function Get-ItemPropertyValue {
    param($LiteralPath, $Name, $ErrorAction)
    Assert-Equal $LiteralPath 'HKLM:\SOFTWARE\Microsoft\RDMonitoringAgent' 'Registry path'
    Assert-Equal $Name 'SessionHostPool' 'Registry value'
    if ($script:registryError) { throw $script:registryError }
    return $script:actualPool
}

$examples = @(
    @('AVDAZINT01A1001', 'P-ARMY-AVD-AZ-INT01A-1'),
    @('AVDAZINT01B2001', 'P-ARMY-AVD-AZ-INT01B-2'),
    @('AVDVAINT02A1001', 'P-ARMY-AVD-VA-INT02A-1'),
    @('AVDVAINT01B2001', 'P-ARMY-AVD-VA-INT01B-2'),
    @('AVDAZPLT03A2001', 'P-ARMY-AVD-AZ-PLT03A-2'),
    @('AVDAZPLT03B1001', 'P-ARMY-AVD-AZ-PLT03B-1'),
    @('AVDAZBRD04A1001', 'P-ARMY-AVD-AZ-BRD04A-1'),
    @('AVDAZBRD04B2001', 'P-ARMY-AVD-AZ-BRD04B-2'),
    @('AVDTXPLT02A1001', 'P-ARMY-AVD-AZ-PLT02A-1'),
    @('AVDTXPLT02B2001', 'P-ARMY-AVD-AZ-PLT02B-2'),
    @('AVDTXBRD03A2001', 'P-ARMY-AVD-AZ-BRD03A-2'),
    @('AVDTXBRD03B1001', 'P-ARMY-AVD-AZ-BRD03B-1'),
    @('AVDDEPLT02A1001', 'P-ARMY-AVD-VA-PLT02A-1'),
    @('AVDDEBRD03B2001', 'P-ARMY-AVD-VA-BRD03B-2'),
    @('AVDDEBRD03A1001', 'P-ARMY-AVD-VA-BRD03A-1'),
    @('AVDDEPLT02B2001', 'P-ARMY-AVD-VA-PLT02B-2'),
    @('AVDVATST01A1001', 'P-ARMY-AVD-VA-TST01A-1'),
    @('AVDVATST01B2001', 'P-ARMY-AVD-VA-TST01B-2')
)
foreach ($example in $examples) {
    $script:actualPool = $example[1]
    $result = Get-HostPoolResult -DeviceName $example[0]
    Assert-Equal $result.Result 'Pass' "Mapping $($example[0])"
    if ($result.Evidence -notlike "*Device=$($example[0])*Actual pool=$($example[1])*Expected pool=$($example[1])*") {
        throw "Incomplete evidence: $($result.Evidence)"
    }
}

$mismatches = @(
    @('P-ARMY-AVD-AZ-INT01B-2', 'host-pool number'),
    @('P-ARMY-AVD-AZ-INT01A-1', 'slice'),
    @('P-ARMY-AVD-AZ-INT02B-1', 'group'),
    @('P-ARMY-AVD-VA-INT01B-1', 'region'),
    @('P-ARMY-AVD-AZ-PLT01B-1', 'ring'),
    @('prefix-P-ARMY-AVD-AZ-INT01B-1', 'format'),
    @('P-ARMY-AVD-AZ-INT01B-1-extra', 'format'),
    @("P-ARMY-AVD-AZ-INT01B-1`n", 'format'),
    @(' P-ARMY-AVD-AZ-INT01B-1 ', 'format'),
    @('', 'blank'),
    @('   ', 'blank'),
    @($null, 'blank')
)
foreach ($case in $mismatches) {
    $script:actualPool = $case[0]
    $result = Get-HostPoolResult -DeviceName 'AVDAZINT01B1001'
    Assert-Equal $result.Result 'Fail' "Reject pool '$($case[0])'"
    if ($result.Evidence -notlike "*$($case[1])*") { throw "Missing mismatch reason '$($case[1])': $($result.Evidence)" }
    if (-not $result.Remediation) { throw 'Missing failure remediation.' }
}
foreach ($device in @('', 'UnknownVM', 'AVDNYINT01B1001', 'AVDAZDEV01B1001', 'AVDAZINT1B1001', 'AVDAZINT01C1001', 'AVDAZINT01B001', 'AVDAZINT01B0001', 'AVDAZINT01BP1001', "AVDAZINT01B1001`n")) {
    $script:actualPool = 'P-ARMY-AVD-AZ-INT01B-1'
    $result = Get-HostPoolResult -DeviceName $device
    Assert-Equal $result.Result 'Fail' "Reject device '$device'"
    if ($result.Evidence -notlike '*device naming*') { throw "Missing device naming reason: $($result.Evidence)" }
}
foreach ($errorMessage in @('Registry key is missing', 'SessionHostPool property is missing', 'Access denied')) {
    $script:registryError = $errorMessage
    $result = Get-HostPoolResult -DeviceName 'AVDAZINT01B1001'
    Assert-Equal $result.Result 'Fail' $errorMessage
    if ($result.Evidence -notlike "*$errorMessage*") { throw "Registry error was not reported: $($result.Evidence)" }
}
$script:registryError = ''
$script:actualPool = 'p-army-avd-az-int01b-1'
Assert-Equal (Get-HostPoolResult -DeviceName 'avdazint01b1001').Result 'Pass' 'Case-insensitive Windows naming'
foreach ($case in @(
    @('AVDDEPLT02A1001', 'P-ARMY-AVD-AZ-PLT02A-1'),
    @('AVDTXPLT02A1001', 'P-ARMY-AVD-VA-PLT02A-1')
)) {
    $script:actualPool = $case[1]
    $result = Get-HostPoolResult -DeviceName $case[0]
    Assert-Equal $result.Result 'Fail' "Reject cross-region pool for $($case[0])"
    if ($result.Evidence -notlike '*region mismatch*') { throw "Missing cross-region failure: $($result.Evidence)" }
}
$script:actualPool = 123
Assert-Equal (Get-HostPoolResult -DeviceName 'AVDAZINT01B1001').Result 'Fail' 'Reject non-string registry value'
Write-Output "PASS: $script:assertions assertions; all 18 supplied mappings, mismatch dimensions, malformed inputs and registry failures."
