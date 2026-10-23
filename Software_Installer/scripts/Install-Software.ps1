#requires -Version 5.1
[CmdletBinding()]
param(
    [string[]]$Include = @(),
    [string[]]$Exclude = @(),
    [switch]$ListOnly,
    [switch]$WhatIf,
    [switch]$NoAutoElevate
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$installerPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'Install-Software.ps1'
if (-not (Test-Path -LiteralPath $installerPath -PathType Leaf)) {
    throw "Canonical installer not found: $installerPath"
}

$forwardedParameters = @{}
if ($Include.Count -gt 0) { $forwardedParameters.Include = $Include }
if ($Exclude.Count -gt 0) { $forwardedParameters.Exclude = $Exclude }
if ($ListOnly) { $forwardedParameters.ListOnly = $true }
if ($WhatIf) { $forwardedParameters.WhatIf = $true }
if ($NoAutoElevate) { $forwardedParameters.NoAutoElevate = $true }

& $installerPath @forwardedParameters
exit $LASTEXITCODE

