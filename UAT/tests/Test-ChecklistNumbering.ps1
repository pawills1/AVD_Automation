[CmdletBinding()]
param(
    [string]$EnginePath,
    [string]$WorkbookPath
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$uatDirectory = Split-Path -Parent $PSScriptRoot
if (-not $EnginePath) { $EnginePath = Join-Path $uatDirectory 'AVD-UAT-Checklist.ps1' }
if (-not $WorkbookPath) { $WorkbookPath = Join-Path $uatDirectory 'AVD Test Plan Template - DO NOT DELETE OR MOVE.xlsx' }
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
$routing = @($definitions | Where-Object Name -eq 'Get-AutomatedResult')[0]
$switches = @($routing.Body.FindAll({ param($node) $node -is [System.Management.Automation.Language.SwitchStatementAst] }, $false))
if ($switches.Count -ne 2) { throw 'Expected separate Remote Apps and Azure Virtual Desktop routing switches.' }

$expectedMarkers = @'
Get-PerformanceResult
Microsoft Graph
Get-SevenZipResult
licensing prompt
Complete the checklist activities
Adobe Acrobat sensitivity-label
AWSWickrGov
Rendered banner
DaxStudio
DEPTH_Bios
Get-DtpLiteResult
Get-FoundryOdbcResult
interactive Chrome
Get-ChromePurviewResult
Get-ChromeUpdatePolicyResult
Get-ChromeManagedBookmarksResult
AIModeSettings
HistorySearchSettings
Get-CopilotArtifactsResult
Get-OfficeStartupResult
Get-OfficeUpdatePolicyResult
Teams media optimization
Get-TeamsStartupResult
Pop out chat
microphone and speaker selectors
Microsoft Teams functionality
interactive Firefox
Get-FirefoxProxyPolicyResult
Get-FirefoxUpdatePolicyResult
Get-EdgePurviewResult
two managed Edge cycles
Get-EdgeDeveloperToolsResult
Get-EdgeDownloadLocationResult
Get-EdgeManagedFavoritesResult
Get-EdgeHomePageResult
Get-EdgeInPrivateResult
encrypted message
Get-EdgeWritingAssistanceResult
notepad++
C:\ocetool\oce.exe
Get-OceShortcutResult
Outlook Cached Exchange Mode
OUTLOOK.EXE
Get-PowerAutomateUpdateResult
PBIDesktop.exe
PowerBIReportBuilder.exe
Get-PowerShellSevenResult
Get-GraphIntunePolicyResult
PrinterLogic
Purebred
SAPlogon.exe
Open SelfHelp
test favorite or bookmark
OPSEC
both documented snip tests
black or blank
STIG Viewer 3.exe
Get-YubiKeyDetectionResult
IDCO CAC
Army wallpaper
Get-BitLockerProtectionResult
C:\ drive permissions
D:\ drive permissions
Get-MicrosoftStoreBlockResult
Get-WindowsAppsStartupResult
Get-InstallerBlockResult
Get-ClipchampBlockResult
Get-InstalledLanguageResult
Get-PuttyDownloadBlockResult
OneDrive accessibility
Always keep on this device
Get-OsBuildResult
Get-LocalUser
Get-WindowsUpdateUiPolicyResult
Get-WindowsUpdateAdvancedOptionsResult
prdweb.afmiscloud.army.mil
atrrs.army.mil
Get-FmswebResult
gears.army.mil
Get-BastionRdpResult
ECMA-WCF-TRUST
Get-ProtectedRootsResult
'@ -split '\r?\n'
if ($expectedMarkers.Count -ne 82) { throw 'Expected markers must cover AVD-00 through AVD-81.' }
$avdClauses = @($switches[1].Clauses)
if ($avdClauses.Count -ne $expectedMarkers.Count) { throw 'Unexpected Azure Virtual Desktop route count.' }
foreach ($number in 0..81) {
    $id = 'AVD-{0:D2}' -f $number
    $clauses = @($avdClauses | Where-Object { $_.Item1.Value -eq $id })
    if ($clauses.Count -ne 1 -or -not $clauses[0].Item2.Extent.Text.Contains($expectedMarkers[$number])) {
        throw "$id must route to '$($expectedMarkers[$number])'."
    }
}

$sourceCsv = Join-Path $env:TEMP ('AVD-Numbering-' + [guid]::NewGuid().ToString('N') + '.csv')
try {
    & (Join-Path $uatDirectory 'tools\AvdChecklistTool\AvdChecklistTool.exe') export --input $WorkbookPath --output $sourceCsv
    if ($LASTEXITCODE -ne 0) { throw "Workbook export failed with exit code $LASTEXITCODE." }
    $checks = @(Import-Csv -LiteralPath $sourceCsv)
    foreach ($mapping in @(
        @{ Sheet = 'Remote Apps Tasks'; Switch = $switches[0]; Count = 17 },
        @{ Sheet = 'Azure Virtual Desktop Tasks'; Switch = $switches[1]; Count = 82 }
    )) {
        $rows = @($checks | Where-Object Sheet -eq $mapping.Sheet)
        $ids = @($mapping.Switch.Clauses | ForEach-Object { $_.Item1.Value })
        if ($rows.Count -ne $mapping.Count -or $ids.Count -ne $mapping.Count) { throw "Unexpected check count for $($mapping.Sheet)." }
        if (@($rows.ID | Select-Object -Unique).Count -ne $rows.Count) { throw "Duplicate workbook IDs in $($mapping.Sheet)." }
        if (@(Compare-Object $rows.ID $ids).Count -ne 0) { throw "Workbook and engine IDs differ for $($mapping.Sheet)." }
    }
    foreach ($expected in @(
        @{ ID = 'AVD-05'; Test = 'Adobe Acrobat 64-Bit' },
        @{ ID = 'AVD-06'; Test = 'AWSWickrGov' },
        @{ ID = 'AVD-80'; Test = 'Certificate Verification' },
        @{ ID = 'AVD-81'; Test = 'Protected Roots Verification' }
    )) {
        $row = @($checks | Where-Object { $_.Sheet -eq 'Azure Virtual Desktop Tasks' -and $_.ID -eq $expected.ID })
        if ($row[0].Test -ne $expected.Test) { throw "Workbook $($expected.ID) must be '$($expected.Test)'." }
    }
}
finally {
    if (Test-Path -LiteralPath $sourceCsv) { Remove-Item -LiteralPath $sourceCsv }
}

function Get-ApplicationStartupResult {
    param([string]$Path, [string]$ProcessName, [string]$DisplayName, [string]$Guidance)
    if ($ProcessName -ne 'AWSWickrGov' -or $DisplayName -ne 'AWS WickrGov') { throw 'AVD-06 routed to the wrong application.' }
    return New-Result 'Partial Pass' $Path $Guidance
}
foreach ($isSystem in @($true, $false)) {
    foreach ($RunInteractiveChecks in @($false, $true)) {
        $result = Get-AutomatedResult 'Azure Virtual Desktop Tasks' 'AVD-05'
        if ($result.Result -ne 'Manual' -or $result.Review -notlike 'Manual task:*') { throw 'AVD-05 must always be manual.' }
        $description = Get-ValidationCheckDescription 'Azure Virtual Desktop Tasks' 'AVD-05' ''
        foreach ($text in @('Adobe Acrobat', 'Microsoft Sensitivity Label', 'save, close, and reopen')) {
            if (-not $result.Notes.Contains($text) -or -not $description.Contains($text)) { throw "AVD-05 guidance is missing '$text'." }
        }
        if ((Get-AutomatedResult 'Azure Virtual Desktop Tasks' 'AVD-06').Result -ne 'Partial Pass') { throw 'AVD-06 startup routing changed.' }
        $remote = Get-AutomatedResult 'Remote Apps Tasks' 'AVD-05'
        if ($remote.Result -ne 'Manual' -or $remote.Notes -notlike '*Microsoft Teams*') { throw 'Remote Apps AVD-05 must remain Teams.' }
    }
}
if ($null -ne (Get-AutomatedResult 'Azure Virtual Desktop Tasks' 'AVD-82')) { throw 'Unexpected route beyond the current workbook.' }
Write-Output 'PASS: all 82 AVD routes, 99 worksheet-scoped workbook IDs, manual Adobe workflow, WickrGov routing, and unchanged Remote Apps AVD-05 verified without live probes.'
