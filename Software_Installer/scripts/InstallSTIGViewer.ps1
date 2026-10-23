#requires -Version 5.1
[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$softwareRoot = Split-Path -Parent $PSScriptRoot
$sourceDirectory = Join-Path $softwareRoot 'U_STIGViewer-win32_x64-3-8-0'
$sourceExecutable = Join-Path $sourceDirectory 'STIG Viewer 3.exe'
$destinationDirectory = Join-Path $env:ProgramFiles 'STIG Viewer 3'
$destinationExecutable = Join-Path $destinationDirectory 'STIG Viewer 3.exe'

if (-not (Test-Path -LiteralPath $sourceExecutable -PathType Leaf)) {
    throw "STIG Viewer executable not found: $sourceExecutable"
}

New-Item -Path $destinationDirectory -ItemType Directory -Force | Out-Null
Copy-Item -Path (Join-Path $sourceDirectory '*') -Destination $destinationDirectory -Recurse -Force

$shell = New-Object -ComObject WScript.Shell
foreach ($shortcutPath in @(
    'C:\Users\Public\Desktop\STIG Viewer 3.lnk'
    'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\STIG Viewer 3.lnk'
)) {
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $destinationExecutable
    $shortcut.WorkingDirectory = $destinationDirectory
    $shortcut.IconLocation = "$destinationExecutable,0"
    $shortcut.Save()
}

if (-not (Test-Path -LiteralPath $destinationExecutable -PathType Leaf)) {
    throw "STIG Viewer deployment verification failed: $destinationExecutable"
}

Write-Output "STIG Viewer 3 deployed successfully to '$destinationDirectory'."
