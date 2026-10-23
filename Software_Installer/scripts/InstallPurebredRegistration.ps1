#requires -Version 5.1
[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$softwareRoot = Split-Path -Parent $PSScriptRoot
$workDir = Join-Path $softwareRoot 'PurebredRegistration'
$bundle = @(Get-ChildItem -LiteralPath $workDir -Filter '*.appxbundle' -File)

if ($bundle.Count -ne 1) {
    throw "Expected exactly one Purebred app bundle in '$workDir'; found $($bundle.Count)."
}

$dependencyNames = @(
    'Microsoft.NET.Native.Framework.1.3.appx'
    'Microsoft.NET.Native.Runtime.1.4.appx'
    'Microsoft.VCLibs.x64.14.00.appx'
)
$dependencyPaths = @($dependencyNames | ForEach-Object {
    $path = Join-Path $workDir "Dependencies\x64\$_"
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Purebred dependency not found: $path"
    }
    $path
})

Write-Output "Provisioning $($bundle[0].Name) for all users..."
$provisionedPackage = Add-AppxProvisionedPackage -Online -SkipLicense -PackagePath $bundle[0].FullName -DependencyPackagePath $dependencyPaths

if ([string]::IsNullOrWhiteSpace($provisionedPackage.PackageName)) {
    throw 'Purebred provisioning completed without returning a package identity.'
}

$verifiedPackage = Get-AppxProvisionedPackage -Online | Where-Object { $_.PackageName -eq $provisionedPackage.PackageName }
if ($null -eq $verifiedPackage) {
    throw "Purebred package '$($provisionedPackage.PackageName)' was not found after provisioning."
}

Write-Output "Purebred Registration provisioned successfully as '$($provisionedPackage.PackageName)'."

