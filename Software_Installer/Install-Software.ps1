#requires -Version 5.1
<#
.SYNOPSIS
Installs the software packages stored in the parent Software directory.

.DESCRIPTION
Runs each enabled package silently, records detailed logs, displays installation
progress, and prints a final success/failure summary. Run -ListOnly before the
first deployment to review the package paths and vendor-specific silent switches.

.EXAMPLE
.\Install-Software.ps1

.EXAMPLE
.\Install-Software.ps1 -WhatIf

.EXAMPLE
.\Install-Software.ps1 -Include '7-Zip','Mozilla Firefox'
#>
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

$softwareRoot = if (Test-Path -LiteralPath (Join-Path $PSScriptRoot '7z2603-x64.exe') -PathType Leaf) {
    $PSScriptRoot
}
else {
    Split-Path -Parent $PSScriptRoot
}
$logRoot = Join-Path $PSScriptRoot 'Logs'
$runId = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
$runLogDirectory = Join-Path $logRoot $runId
$summaryPath = Join-Path $runLogDirectory 'Install-Summary.csv'
$transcriptPath = Join-Path $runLogDirectory 'Install-Transcript.log'
$successExitCodes = @(0, 1641, 3010)

# Package arguments are vendor-specific. Keep them here so they can be reviewed
# and updated without changing the installation engine below.
$packages = @(
    [pscustomobject]@{ Name = '7-Zip'; Path = '7z2603-x64.exe'; Type = 'Exe'; Arguments = '/S'; Enabled = $true; Comment = '7-Zip silent installer.' },
    [pscustomobject]@{ Name = 'AWSWickrGov'; Path = 'AWSWickrGov'; Type = 'AutoInstaller'; Arguments = '/s'; Enabled = $true; Comment = 'Installs the single EXE or MSI found in the AWSWickrGov folder (including ZIP contents).' },
    [pscustomobject]@{ Name = 'Google Chrome'; Path = 'ChromeSetup.exe'; Type = 'Exe'; Arguments = '/silent /install'; Enabled = $true; Comment = 'Google installer silent mode; may require network access if this is a web stub.' },
    [pscustomobject]@{ Name = 'DAX Studio'; Path = 'DaxStudio_3_6_1_setup.exe'; Type = 'Exe'; Arguments = '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-'; Enabled = $true; Comment = 'Inno Setup silent mode.' },
    [pscustomobject]@{ Name = 'Mozilla Firefox'; Path = 'Firefox Installer.exe'; Type = 'Exe'; Arguments = '-ms'; Enabled = $true; Comment = 'Mozilla silent mode; may require network access if this is a web stub.' },
    [pscustomobject]@{ Name = 'Foundry SQL ODBC Driver'; Path = 'foundry-sql-odbc-driver-3.41.0.exe'; Type = 'Exe'; Arguments = '/quiet /norestart'; Enabled = $true; Comment = 'Verify the vendor silent switches in a test image.' },
    [pscustomobject]@{ Name = 'GCSS-Army Integrated Installer'; Path = 'GCSS-ArmyIntegratedInstaller5.7.0.exe'; Type = 'Exe'; Arguments = '/quiet /norestart'; Enabled = $true; Comment = 'Verify the vendor silent switches in a test image.' },
    [pscustomobject]@{ Name = 'InstallRoot'; Path = 'InstallRoot_5.6x64.msi'; Type = 'Msi'; Arguments = ''; Enabled = $true; Comment = 'MSI quiet installation.' },
    [pscustomobject]@{ Name = 'Remote Desktop WebRTC Redirector'; Path = 'MsRdcWebRTCSvc_HostSetup_1.50.2402.29001_x64.msi'; Type = 'Msi'; Arguments = ''; Enabled = $true; Comment = 'MSI quiet installation.' },
    [pscustomobject]@{ Name = 'NetBanner'; Path = 'NetBanner.Setup.msi'; Type = 'Msi'; Arguments = ''; Enabled = $true; Comment = 'MSI quiet installation.' },
    [pscustomobject]@{ Name = 'Power BI Desktop'; Path = 'PBIDesktopSetup_x64.exe'; Type = 'Exe'; Arguments = '-quiet -norestart ACCEPT_EULA=1'; Enabled = $true; Comment = 'Microsoft Power BI documented silent mode.' },
    [pscustomobject]@{ Name = 'Power BI Report Builder'; Path = 'PowerBIReportBuilder.msi'; Type = 'Msi'; Arguments = 'ACCEPT_EULA=1'; Enabled = $true; Comment = 'MSI quiet installation with EULA acceptance.' },
    [pscustomobject]@{ Name = 'Printer Installer Client'; Path = 'PrinterInstallerClient.msi'; Type = 'Msi'; Arguments = ''; Enabled = $true; Comment = 'MSI quiet installation.' },
    [pscustomobject]@{ Name = 'S/MIME for Outlook Web'; Path = 'SmimeOutlookWebChrome.msi'; Type = 'Msi'; Arguments = ''; Enabled = $true; Comment = 'MSI quiet installation.' },
    [pscustomobject]@{ Name = 'YubiKey Manager'; Path = 'yubikey-manager-5.9.2-win64.msi'; Type = 'Msi'; Arguments = ''; Enabled = $true; Comment = 'MSI quiet installation.' },
    [pscustomobject]@{ Name = 'Microsoft 365 Apps'; Path = 'M365_Office\OfficeSetup.exe'; Configuration = 'M365_Office\AVDOffice.xml'; Type = 'Office'; Arguments = ''; Enabled = $true; Comment = 'Installs Microsoft 365 Apps, Visio, and Project using AVDOffice.xml.' },
    [pscustomobject]@{ Name = 'Microsoft Teams'; Path = 'M365_Office\Teams_windows_x64.msi'; Type = 'Msi'; Arguments = 'ALLUSERS=1'; Enabled = $true; Comment = 'Teams per-machine silent MSI installation.' },
    [pscustomobject]@{ Name = 'Microsoft Teams Meeting Add-in'; Path = 'M365_Office\MicrosoftTeamsMeetingAddinInstaller.msi'; Type = 'TeamsMeetingAddIn'; Arguments = ''; Enabled = $true; Comment = 'Installs the Teams meeting add-in using its MSI product version for TARGETDIR.' },
    [pscustomobject]@{ Name = 'Notepad++'; Path = 'npp*.exe'; Type = 'NotepadPlusPlus'; Arguments = '/S'; Enabled = $true; Comment = 'Installs the single npp*.exe offline installer found in the software directory.' },
    [pscustomobject]@{ Name = 'OCE Tool'; Path = 'OCE.EXE'; Shortcut = 'OCE Tool.lnk'; Type = 'PortableWithShortcuts'; Arguments = ''; Enabled = $true; Comment = 'Deploys OCE to C:\OCETOOL and creates shared desktop and Start menu shortcuts.' },
    [pscustomobject]@{ Name = 'STIG Viewer 3'; Path = 'U_STIGViewer-win32_x64-3-8-0'; Type = 'STIGViewer'; Arguments = ''; Enabled = $true; Comment = 'Deploys the unpacked STIG Viewer application and all-users shortcuts.' },
    [pscustomobject]@{ Name = 'Self-Help'; Path = 'scripts\InstallSelfHelp.ps1'; Type = 'PowerShell'; Arguments = ('-Install -WorkDir "{0}"' -f (Join-Path $softwareRoot 'SelfHelp')); Enabled = $true; Comment = 'Runs InstallSelfHelp.ps1 to deploy Self-Help and configure its JEA endpoint.' },
    [pscustomobject]@{ Name = 'Army Wallpaper'; Path = 'Army.png'; Type = 'Wallpaper'; Arguments = ''; Enabled = $true; Comment = 'Adds Army.png to the Windows wallpaper collection for all users.' },
    [pscustomobject]@{ Name = 'FSLogix Apps'; Path = 'FSLogix_25.04\x64\Release\FSLogixAppsSetup.exe'; Type = 'Exe'; Arguments = '/install /quiet /norestart'; Enabled = $true; Comment = 'FSLogix core product.' },
    [pscustomobject]@{ Name = 'FSLogix Rule Editor'; Path = 'FSLogix_25.04\x64\Release\FSLogixAppsRuleEditorSetup.exe'; Type = 'Exe'; Arguments = '/install /quiet /norestart'; Enabled = $true; Comment = 'FSLogix administration tool.' },
    [pscustomobject]@{ Name = 'Purebred Registration'; Path = 'scripts\InstallPurebredRegistration.ps1'; Type = 'PowerShell'; Arguments = ''; Enabled = $true; Comment = 'Runs InstallPurebredRegistration.ps1 to provision the app bundle and x64 dependencies for all users, then verifies the package.' },
    [pscustomobject]@{ Name = 'Tychon'; Path = 'Tychon_AVD\Tychon.bat'; Type = 'Cmd'; Arguments = ''; Enabled = $true; Comment = 'Runs the package-provided compatibility installer.' },
    [pscustomobject]@{ Name = 'Adobe Acrobat Professional'; Path = 'Unified_Installer.ps1'; Type = 'PowerShell'; Arguments = '-version Acrobatx64'; Enabled = $true; Comment = 'Runs the supplied Acrobat Pro install, patch, licensing, and STIG workflow.' },
    [pscustomobject]@{ Name = 'DISA DEPTH'; Path = 'DEPTH\Source\install.ps1'; Type = 'PowerShell'; Arguments = ''; Enabled = $true; Comment = 'Installs the DEPTH module and scheduled task using the supplied deployment script.' },
    [pscustomobject]@{ Name = 'DTP Lite'; Path = 'DTP_Lite_v6.0.8.2\DTP_Lite.ps1'; Type = 'PowerShell'; Arguments = ''; Enabled = $true; Comment = 'Copies DTP Lite into ProgramData and creates shared shortcuts.' },
    [pscustomobject]@{ Name = 'IADS4'; Path = 'IADS4_Workspace\IADS_Setup.exe'; Type = 'Exe'; Arguments = '/S'; Enabled = $true; Comment = 'Uses the vendor-documented silent multi-user installation.' }
)

# These parent-folder items are intentionally not counted as installers.
$notInstallers = @(
    'Acrobat_Detection.ps1, PrinterDevMode_Task.ps1, and registry_fix.txt (support files used by Adobe deployment)',
    'VDOT_AVD (system optimization tool; run separately after testing)'
)

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Start-ElevatedCopy {
    $forwardedArguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $PSCommandPath))
    if ($Include.Count -gt 0) {
        $includeValues = ($Include | ForEach-Object { '"{0}"' -f ($_ -replace '"', '`"') }) -join ','
        $forwardedArguments += @('-Include', $includeValues)
    }
    if ($Exclude.Count -gt 0) {
        $excludeValues = ($Exclude | ForEach-Object { '"{0}"' -f ($_ -replace '"', '`"') }) -join ','
        $forwardedArguments += @('-Exclude', $excludeValues)
    }
    if ($ListOnly) { $forwardedArguments += '-ListOnly' }
    if ($WhatIf) { $forwardedArguments += '-WhatIf' }
    $process = Start-Process -FilePath 'powershell.exe' -ArgumentList $forwardedArguments -Verb RunAs -Wait -PassThru
    exit $process.ExitCode
}

function Write-Status {
    param([string]$Message, [ConsoleColor]$Color = [ConsoleColor]::Gray)
    $timestampedMessage = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Write-Host $timestampedMessage -ForegroundColor $Color
}

function Get-SafeFileName {
    param([string]$Name)
    return ($Name -replace '[^A-Za-z0-9._-]', '_')
}

function Get-MsiProperty {
    param(
        [string]$Path,
        [string]$Property
    )

    $installer = New-Object -ComObject WindowsInstaller.Installer
    $database = $installer.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $installer, @($Path, 0))
    $view = $database.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $database, @("SELECT `Value` FROM `Property` WHERE `Property`='$Property'"))
    $view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null) | Out-Null
    $record = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)
    if ($null -eq $record) {
        throw "MSI property '$Property' was not found in '$Path'."
    }
    return $record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, 1)
}

function Install-NotepadPlusPlusPackage {
    param([string]$SoftwareRoot, [string]$Arguments)

    $installFiles = @(Get-ChildItem -LiteralPath $SoftwareRoot -Filter 'npp*.exe' -File -Recurse)
    if ($installFiles.Count -ne 1) {
        throw "Expected exactly one Notepad++ installer named 'npp*.exe' under '$SoftwareRoot'; found $($installFiles.Count)."
    }
    $process = Start-Process -FilePath $installFiles[0].FullName -ArgumentList $Arguments -Wait -PassThru
    if ($process.ExitCode -ne 0) {
        return $process.ExitCode
    }
    $installedExecutable = Join-Path $env:ProgramFiles 'Notepad++\notepad++.exe'
    if (-not (Test-Path -LiteralPath $installedExecutable -PathType Leaf)) {
        throw "Notepad++ verification failed: $installedExecutable was not found."
    }
    return 0
}

function Install-STIGViewerPackage {
    param([string]$PackagePath)

    $sourceExecutable = Join-Path $PackagePath 'STIG Viewer 3.exe'
    if (-not (Test-Path -LiteralPath $sourceExecutable -PathType Leaf)) {
        throw "STIG Viewer executable not found: $sourceExecutable"
    }
    $destinationDirectory = Join-Path $env:ProgramFiles 'STIG Viewer 3'
    $destinationExecutable = Join-Path $destinationDirectory 'STIG Viewer 3.exe'
    New-Item -Path $destinationDirectory -ItemType Directory -Force | Out-Null
    Copy-Item -Path (Join-Path $PackagePath '*') -Destination $destinationDirectory -Recurse -Force
    $shell = New-Object -ComObject WScript.Shell
    foreach ($shortcutPath in @('C:\Users\Public\Desktop\STIG Viewer 3.lnk', 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\STIG Viewer 3.lnk')) {
        $shortcut = $shell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $destinationExecutable
        $shortcut.WorkingDirectory = $destinationDirectory
        $shortcut.IconLocation = "$destinationExecutable,0"
        $shortcut.Save()
    }
    foreach ($requiredPath in @($destinationExecutable, 'C:\Users\Public\Desktop\STIG Viewer 3.lnk', 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\STIG Viewer 3.lnk')) {
        if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
            throw "STIG Viewer verification failed: $requiredPath was not found."
        }
    }
    return 0
}

function Install-PurebredPackage {
    param([string]$PackagePath)

    $bundles = @(Get-ChildItem -LiteralPath $PackagePath -Filter '*.appxbundle' -File)
    if ($bundles.Count -ne 1) {
        throw "Expected exactly one Purebred app bundle in '$PackagePath'; found $($bundles.Count)."
    }
    $dependencyPaths = @('Microsoft.NET.Native.Framework.1.3.appx', 'Microsoft.NET.Native.Runtime.1.4.appx', 'Microsoft.VCLibs.x64.14.00.appx') | ForEach-Object {
        $dependencyPath = Join-Path $PackagePath "Dependencies\x64\$_"
        if (-not (Test-Path -LiteralPath $dependencyPath -PathType Leaf)) {
            throw "Purebred dependency not found: $dependencyPath"
        }
        $dependencyPath
    }
    $provisionedPackage = Add-AppxProvisionedPackage -Online -SkipLicense -PackagePath $bundles[0].FullName -DependencyPackagePath $dependencyPaths
    if ([string]::IsNullOrWhiteSpace($provisionedPackage.PackageName)) {
        throw 'Purebred provisioning completed without returning a package identity.'
    }
    $verifiedPackage = Get-AppxProvisionedPackage -Online | Where-Object { $_.PackageName -eq $provisionedPackage.PackageName } | Select-Object -First 1
    if ($null -eq $verifiedPackage) {
        throw "Purebred verification failed: package '$($provisionedPackage.PackageName)' was not found."
    }
    return 0
}

function Invoke-PackageInstall {
    param(
        [pscustomobject]$Package,
        [int]$Number,
        [int]$Total
    )

    $packagePath = Join-Path $softwareRoot $Package.Path
    $safeName = Get-SafeFileName -Name $Package.Name
    $detailLog = Join-Path $runLogDirectory ($safeName + '.log')
    $stdoutLog = Join-Path $runLogDirectory ($safeName + '-stdout.log')
    $stderrLog = Join-Path $runLogDirectory ($safeName + '-stderr.log')
    $started = Get-Date
    $result = 'Failed'
    $exitCode = $null
    $message = ''
    $verificationCompleted = $false

    Write-Progress -Activity 'Installing software' -Status ("{0} of {1}: {2}" -f $Number, $Total, $Package.Name) -PercentComplete ((($Number - 1) / $Total) * 100)
    Write-Status ("[{0}/{1}] Starting {2}" -f $Number, $Total, $Package.Name) Cyan

    $requiredPathType = if ($Package.Type -eq 'NotepadPlusPlus') { 'Pattern' } elseif ($Package.Type -in @('AutoInstaller', 'STIGViewer', 'Purebred')) { 'Container' } else { 'Leaf' }
    $packagePathExists = if ($requiredPathType -eq 'Pattern') {
        @(Get-ChildItem -LiteralPath $softwareRoot -Filter $Package.Path -File -Recurse).Count -eq 1
    }
    else {
        Test-Path -LiteralPath $packagePath -PathType $requiredPathType
    }
    if (-not $packagePathExists) {
        $message = if ($requiredPathType -eq 'Pattern') { "Expected exactly one package matching '$($Package.Path)' under '$softwareRoot'." } else { "Package file not found: $packagePath" }
    }
    elseif ($WhatIf) {
        $result = 'WhatIf'
        $exitCode = 0
        $message = 'Installation was not started.'
    }
    else {
        try {
            switch ($Package.Type) {
                'Msi' {
                    $arguments = '/i "{0}" /qn /norestart /L*v "{1}" {2}' -f $packagePath, $detailLog, $Package.Arguments
                    $process = Start-Process -FilePath 'msiexec.exe' -ArgumentList $arguments -Wait -PassThru
                }
                'Exe' {
                    $process = Start-Process -FilePath $packagePath -ArgumentList $Package.Arguments -WorkingDirectory (Split-Path -Parent $packagePath) -RedirectStandardOutput $stdoutLog -RedirectStandardError $stderrLog -Wait -PassThru
                }
                'PowerShell' {
                    $arguments = '-NoProfile -ExecutionPolicy Bypass -File "{0}" {1}' -f $packagePath, $Package.Arguments
                    $process = Start-Process -FilePath 'powershell.exe' -ArgumentList $arguments -WorkingDirectory (Split-Path -Parent $packagePath) -RedirectStandardOutput $stdoutLog -RedirectStandardError $stderrLog -Wait -PassThru
                }
                'Cmd' {
                    $arguments = '/d /c ""{0}" {1}"' -f $packagePath, $Package.Arguments
                    $process = Start-Process -FilePath 'cmd.exe' -ArgumentList $arguments -WorkingDirectory (Split-Path -Parent $packagePath) -RedirectStandardOutput $stdoutLog -RedirectStandardError $stderrLog -Wait -PassThru
                }
                'Office' {
                    $configurationPath = Join-Path $softwareRoot $Package.Configuration
                    if (-not (Test-Path -LiteralPath $configurationPath -PathType Leaf)) {
                        throw "Office configuration file not found: $configurationPath"
                    }
                    [xml]$officeConfiguration = Get-Content -LiteralPath $configurationPath -Raw
                    $productIds = @($officeConfiguration.Configuration.Add.Product | ForEach-Object { $_.ID })
                    foreach ($requiredProductId in @('VisioProRetail', 'ProjectProRetail')) {
                        if ($productIds -notcontains $requiredProductId) {
                            throw "Office configuration must include product '$requiredProductId'."
                        }
                    }
                    $arguments = '/configure "{0}"' -f $configurationPath
                    $process = Start-Process -FilePath $packagePath -ArgumentList $arguments -WorkingDirectory (Split-Path -Parent $packagePath) -Wait -PassThru
                }
                'TeamsMeetingAddIn' {
                    $addInVersion = Get-MsiProperty -Path $packagePath -Property 'ProductVersion'
                    $targetDirectory = "C:\Program Files (x86)\Microsoft\TeamsMeetingAddin\$addInVersion\"
                    $arguments = '/i "{0}" /passive /norestart ALLUSERS=1 TARGETDIR="{1}" /L*v "{2}"' -f $packagePath, $targetDirectory, $detailLog
                    $process = Start-Process -FilePath 'msiexec.exe' -ArgumentList $arguments -Wait -PassThru
                }
                'PortableWithShortcuts' {
                    $targetDirectory = 'C:\OCETOOL'
                    New-Item -Path $targetDirectory -ItemType Directory -Force | Out-Null
                    Copy-Item -LiteralPath $packagePath -Destination (Join-Path $targetDirectory 'OCE.EXE') -Force

                    $shortcutPath = Join-Path $softwareRoot $Package.Shortcut
                    if (-not (Test-Path -LiteralPath $shortcutPath -PathType Leaf)) {
                        throw "Shortcut file not found: $shortcutPath"
                    }
                    Copy-Item -LiteralPath $shortcutPath -Destination 'C:\Users\Public\Desktop\OCE Tool.lnk' -Force
                    Copy-Item -LiteralPath $shortcutPath -Destination 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OCE Tool.lnk' -Force
                    $process = [pscustomobject]@{ ExitCode = 0 }
                }
                'Wallpaper' {
                    $wallpaperDirectory = 'C:\Windows\Web\Wallpaper\Windows'
                    New-Item -Path $wallpaperDirectory -ItemType Directory -Force | Out-Null
                    Copy-Item -LiteralPath $packagePath -Destination (Join-Path $wallpaperDirectory 'Army.png') -Force
                    $process = [pscustomobject]@{ ExitCode = 0 }
                }
                'AutoInstaller' {
                    $installSearchRoot = $packagePath
                    $installFiles = @(Get-ChildItem -LiteralPath $installSearchRoot -Recurse -File | Where-Object { $_.Extension -in @('.exe', '.msi') })
                    if ($installFiles.Count -eq 0) {
                        $zipFiles = @(Get-ChildItem -LiteralPath $packagePath -File -Filter '*.zip')
                        if ($zipFiles.Count -ne 1) {
                            throw "Expected one AWS WickrGov installer or ZIP in '$packagePath'."
                        }
                        $installSearchRoot = Join-Path $runLogDirectory 'AWSWickrGov-media'
                        Expand-Archive -LiteralPath $zipFiles[0].FullName -DestinationPath $installSearchRoot -Force
                        $installFiles = @(Get-ChildItem -LiteralPath $installSearchRoot -Recurse -File | Where-Object { $_.Extension -in @('.exe', '.msi') })
                    }
                    if ($installFiles.Count -ne 1) {
                        throw "Expected exactly one AWS WickrGov EXE or MSI in '$installSearchRoot'; found $($installFiles.Count)."
                    }
                    $installFile = $installFiles[0]
                    if ($installFile.Extension -eq '.msi') {
                        $arguments = '/i "{0}" /qn /norestart ALLUSERS=1 /L*v "{1}"' -f $installFile.FullName, $detailLog
                        $process = Start-Process -FilePath 'msiexec.exe' -ArgumentList $arguments -Wait -PassThru
                    }
                    else {
                        $process = Start-Process -FilePath $installFile.FullName -ArgumentList $Package.Arguments -WorkingDirectory $installFile.DirectoryName -RedirectStandardOutput $stdoutLog -RedirectStandardError $stderrLog -Wait -PassThru
                    }
                }
                'NotepadPlusPlus' {
                    $process = [pscustomobject]@{ ExitCode = Install-NotepadPlusPlusPackage -SoftwareRoot $softwareRoot -Arguments $Package.Arguments }
                    $verificationCompleted = $process.ExitCode -eq 0
                }
                'STIGViewer' {
                    $process = [pscustomobject]@{ ExitCode = Install-STIGViewerPackage -PackagePath $packagePath }
                    $verificationCompleted = $true
                }
                'Purebred' {
                    $process = [pscustomobject]@{ ExitCode = Install-PurebredPackage -PackagePath $packagePath }
                    $verificationCompleted = $true
                }
                default { throw "Unknown package type '$($Package.Type)'." }
            }

            $exitCode = $process.ExitCode

            # The vendor Tychon batch file prints failures but always exits 0.
            if ($Package.Name -eq 'Tychon' -and $exitCode -eq 0 -and (Test-Path -LiteralPath $stdoutLog)) {
                if (Select-String -Path $stdoutLog -Pattern 'Unable to install|Installation aborted' -Quiet) {
                    $exitCode = 1
                }
            }

            # Some supplied PowerShell deployment scripts report failure text but
            # return a value instead of setting the PowerShell process exit code.
            if ($Package.Type -eq 'PowerShell' -and $exitCode -eq 0) {
                $failurePattern = 'Installation failed|not installed successfully|Unable to install|Unable to continue|Write-Error|Exception:'
                $scriptLogs = @($stdoutLog, $stderrLog) | Where-Object { Test-Path -LiteralPath $_ }
                if ($scriptLogs.Count -gt 0 -and (Select-String -Path $scriptLogs -Pattern $failurePattern -Quiet)) {
                    $exitCode = 1
                }
            }

            if ($successExitCodes -contains $exitCode) {
                $result = if ($exitCode -in @(1641, 3010)) { 'SuccessRebootRequired' } else { 'Success' }
                $message = if ($verificationCompleted) { 'Installer completed and post-install verification passed.' } else { 'Installer completed.' }
            }
            else {
                $message = "Installer returned exit code $exitCode."
            }
        }
        catch {
            $message = $_.Exception.Message
        }
    }

    $duration = [math]::Round(((Get-Date) - $started).TotalSeconds, 1)
    $color = if ($result -like 'Success*') { [ConsoleColor]::Green } elseif ($result -eq 'WhatIf') { [ConsoleColor]::Yellow } else { [ConsoleColor]::Red }
    Write-Status ("[{0}/{1}] {2}: {3} ({4}s) {5}" -f $Number, $Total, $Package.Name, $result, $duration, $message) $color

    return [pscustomobject]@{
        Number = $Number
        Software = $Package.Name
        Result = $result
        ExitCode = $exitCode
        DurationSeconds = $duration
        Package = $packagePath
        Message = $message
    }
}

if (-not $ListOnly -and -not $WhatIf -and -not $NoAutoElevate -and -not (Test-IsAdministrator)) {
    Write-Host 'Administrator rights are required. Requesting elevation...'
    Start-ElevatedCopy
}

$selectedPackages = @($packages | Where-Object { $_.Enabled })
if ($Include.Count -gt 0) {
    $selectedPackages = @($selectedPackages | Where-Object { $Include -contains $_.Name })
}
if ($Exclude.Count -gt 0) {
    $selectedPackages = @($selectedPackages | Where-Object { $Exclude -notcontains $_.Name })
}

Write-Host ''
Write-Host 'Unified Software Installer' -ForegroundColor White
Write-Host ("Software directory : {0}" -f $softwareRoot)
Write-Host ("Packages selected  : {0} of {1}" -f $selectedPackages.Count, $packages.Count)
Write-Host ''

if ($ListOnly) {
    $selectedPackages | Select-Object Name, Type, Enabled, Path, Arguments, Comment | Format-Table -Wrap -AutoSize
    Write-Host 'Not treated as installers:' -ForegroundColor Yellow
    $notInstallers | ForEach-Object { Write-Host ("  - {0}" -f $_) }
    exit 0
}

if ($selectedPackages.Count -eq 0) {
    throw 'No packages matched the Include/Exclude filters.'
}

New-Item -Path $runLogDirectory -ItemType Directory -Force | Out-Null
Start-Transcript -Path $transcriptPath -Force | Out-Null
$results = @()

try {
    Write-Status ("Installation run started. Logs: {0}" -f $runLogDirectory) White
    for ($index = 0; $index -lt $selectedPackages.Count; $index++) {
        $results += Invoke-PackageInstall -Package $selectedPackages[$index] -Number ($index + 1) -Total $selectedPackages.Count
    }
}
finally {
    Write-Progress -Activity 'Installing software' -Completed
    $results | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8

    $succeeded = @($results | Where-Object { $_.Result -like 'Success*' }).Count
    $failed = @($results | Where-Object { $_.Result -eq 'Failed' }).Count
    $simulated = @($results | Where-Object { $_.Result -eq 'WhatIf' }).Count
    $rebootRequired = @($results | Where-Object { $_.Result -eq 'SuccessRebootRequired' }).Count

    Write-Host ''
    Write-Host 'Installation Summary' -ForegroundColor White
    $results | Format-Table Number, Software, Result, ExitCode, DurationSeconds -AutoSize
    Write-Host ("Selected: {0}  Succeeded: {1}  Failed: {2}  WhatIf: {3}  Reboot required: {4}" -f $selectedPackages.Count, $succeeded, $failed, $simulated, $rebootRequired)
    Write-Host ("Summary log: {0}" -f $summaryPath)
    Stop-Transcript | Out-Null
}

if (@($results | Where-Object { $_.Result -eq 'Failed' }).Count -gt 0) {
    exit 1
}
exit 0
