                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     [CmdletBinding()]
param(
    [string]$InputFile,
    [string]$OutputDirectory,
    [string]$WorkbookTool,
    [string]$ConfigurationFile,
    [string]$UatDirectory = 'C:\Program Files\ArmyAVD\UAT',
    [switch]$UseExistingResultsDirectory,
    [ValidateRange(1, 3600)]
    [int]$PerformanceSampleSeconds = 60,
    [ValidateRange(1, 300)]
    [int]$PerformanceSampleIntervalSeconds = 5,
    [ValidateRange(1, 100)]
    [double]$MaximumAverageCpuPercent = 40,
    [ValidateRange(1, 100)]
    [double]$MaximumMemoryUsedPercent = 80,
    [ValidateRange(0.1, 10000)]
    [double]$MaximumDiskLatencyMilliseconds = 20,
    [string]$PurviewExtensionId,
    [string]$EdgePurviewExtensionId,
    [string]$ChromeBookmarksBaselineFile,
    [string]$EdgeFavoritesBaselineFile,
    [string]$CopilotArtifactName,
    [string]$CopilotAppId,
    [string]$CopilotUrl,
    [string]$ExeInstallerTestFile,
    [string]$MsiInstallerTestFile,
    [string[]]$RequiredLanguageTags,
    [object[]]$RequiredLanguages,
    [string]$ApprovedOsBuild,
    [string]$WindowsUpdateHistoryUrl,
    [string]$ExpectedRemoteAppHostPattern,
    [string]$OwaUrl,
    [switch]$RunInteractiveChecks
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
if (-not $ConfigurationFile) {
    $defaultConfigurationFile = Join-Path $PSScriptRoot 'avd-uat.config.json'
    if (Test-Path -LiteralPath $defaultConfigurationFile -PathType Leaf) { $ConfigurationFile = $defaultConfigurationFile }
}
if ($ConfigurationFile) {
    if (-not (Test-Path -LiteralPath $ConfigurationFile -PathType Leaf)) { throw "Configuration file not found: $ConfigurationFile" }
    $ConfigurationFile = (Resolve-Path -LiteralPath $ConfigurationFile).Path
    try { $configuration = Get-Content -LiteralPath $ConfigurationFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop }
    catch { throw "Configuration file is not valid JSON: $($_.Exception.Message)" }
    if ($null -eq $configuration -or $configuration -is [System.Array]) { throw 'Configuration file must contain one JSON object.' }

    $allowedConfigurationNames = @(
        'PurviewExtensionId', 'EdgePurviewExtensionId', 'ChromeBookmarksBaselineFile', 'EdgeFavoritesBaselineFile',
        'CopilotArtifactName', 'CopilotAppId', 'CopilotUrl', 'ExeInstallerTestFile', 'MsiInstallerTestFile',
        'RequiredLanguageTags', 'RequiredLanguages', 'ApprovedOsBuild', 'WindowsUpdateHistoryUrl',
        'ExpectedRemoteAppHostPattern', 'OwaUrl'
    )
    $configurationDirectory = Split-Path -Parent $ConfigurationFile
    foreach ($property in $configuration.PSObject.Properties) {
        $name = [string]$property.Name
        if ($name -notin $allowedConfigurationNames) { throw "Unknown configuration property: $name" }
        if ($PSBoundParameters.ContainsKey($name)) { continue }

        $value = $property.Value
        switch ($name) {
            'PurviewExtensionId' { $PurviewExtensionId = [string]$value }
            'EdgePurviewExtensionId' { $EdgePurviewExtensionId = [string]$value }
            'ChromeBookmarksBaselineFile' {
                $ChromeBookmarksBaselineFile = [string]$value
                if ($ChromeBookmarksBaselineFile -and -not [IO.Path]::IsPathRooted($ChromeBookmarksBaselineFile)) { $ChromeBookmarksBaselineFile = Join-Path $configurationDirectory $ChromeBookmarksBaselineFile }
            }
            'EdgeFavoritesBaselineFile' {
                $EdgeFavoritesBaselineFile = [string]$value
                if ($EdgeFavoritesBaselineFile -and -not [IO.Path]::IsPathRooted($EdgeFavoritesBaselineFile)) { $EdgeFavoritesBaselineFile = Join-Path $configurationDirectory $EdgeFavoritesBaselineFile }
            }
            'CopilotArtifactName' { $CopilotArtifactName = [string]$value }
            'CopilotAppId' { $CopilotAppId = [string]$value }
            'CopilotUrl' { $CopilotUrl = [string]$value }
            'ExeInstallerTestFile' {
                $ExeInstallerTestFile = [string]$value
                if ($ExeInstallerTestFile -and -not [IO.Path]::IsPathRooted($ExeInstallerTestFile)) { $ExeInstallerTestFile = Join-Path $configurationDirectory $ExeInstallerTestFile }
            }
            'MsiInstallerTestFile' {
                $MsiInstallerTestFile = [string]$value
                if ($MsiInstallerTestFile -and -not [IO.Path]::IsPathRooted($MsiInstallerTestFile)) { $MsiInstallerTestFile = Join-Path $configurationDirectory $MsiInstallerTestFile }
            }
            'RequiredLanguageTags' { $RequiredLanguageTags = @($value | ForEach-Object { [string]$_ }) }
            'RequiredLanguages' { $RequiredLanguages = @($value | ForEach-Object { $_ }) }
            'ApprovedOsBuild' { $ApprovedOsBuild = [string]$value }
            'WindowsUpdateHistoryUrl' { $WindowsUpdateHistoryUrl = [string]$value }
            'ExpectedRemoteAppHostPattern' { $ExpectedRemoteAppHostPattern = [string]$value }
            'OwaUrl' { $OwaUrl = [string]$value }
        }
    }
}
function Resolve-ChecklistInputFile {
    param([string]$Path, [string]$SearchDirectory)

    if ($Path -and (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return (Resolve-Path -LiteralPath $Path).Path
    }

    $pattern = if ($Path) {
        if ([IO.Path]::IsPathRooted($Path)) { $Path } else { Join-Path $SearchDirectory $Path }
    }
    else {
        Join-Path $SearchDirectory 'AVD Test Plan*'
    }
    $matches = @(Get-ChildItem -Path $pattern -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike '~$*' -and $_.Extension -in @('.xlsx', '.xlsm') } |
        Sort-Object LastWriteTime -Descending)
    if ($matches.Count -eq 0) {
        throw "No Excel workbook matching '$pattern' was found. Expected an .xlsx or .xlsm file whose name starts with 'AVD Test Plan'."
    }
    if ($matches.Count -gt 1) {
        Write-Verbose "Multiple checklist workbooks matched '$pattern'; using the newest: $($matches[0].FullName)"
    }
    return $matches[0].FullName
}

. (Join-Path $PSScriptRoot 'AVD-UAT-Paths.ps1')
$InputFile = Resolve-ChecklistInputFile -Path $InputFile -SearchDirectory $UatDirectory
if (-not $OutputDirectory) { $OutputDirectory = $UatDirectory }
if (-not $UseExistingResultsDirectory) { $OutputDirectory = New-UatResultsDirectory -RootDirectory $OutputDirectory }
if (-not $WorkbookTool) { $WorkbookTool = Join-Path $PSScriptRoot 'tools\AvdChecklistTool\AvdChecklistTool.exe' }
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
$results = @()
$isSystem = $env:USERNAME -eq 'SYSTEM'

function New-Result {
    param(
        [ValidateSet('Pass', 'Partial Pass', 'Fail', 'Manual')]
        [string]$Result,
        [string]$Notes,
        [string]$Review = 'Automated'
    )
    return @{ Result = $Result; Notes = $Notes; Review = $Review }
}

function Get-ValidationCheckDescription {
    param([string]$Sheet, [string]$Id, [string]$Review)

    $key = "$Sheet|$Id"
    $descriptions = @{
        'Remote Apps Tasks|AVD-02' = 'Launch Microsoft Teams, wait up to 30 seconds for ms-teams.exe, then close only the process started by the check.'
        'Remote Apps Tasks|AVD-08' = 'Read HKLM/HKCU:\SOFTWARE\Policies\Microsoft\Edge\DeveloperToolsAvailability and require DWORD 2.'
        'Remote Apps Tasks|AVD-04' = 'Manual only: confirm the microphone and speaker selections in Microsoft Teams.'
        'Remote Apps Tasks|AVD-05' = 'Manual only: confirm Microsoft Teams functions correctly.'
        'Remote Apps Tasks|AVD-09' = 'Require HKLM Edge PromptForDownloadLocation=1 and discover the Army OneDrive DownloadDirectory Value beneath any HKLM:\SOFTWARE\Microsoft\MDMWins identity.'
        'Remote Apps Tasks|AVD-10' = 'Read ManagedFavorites from the approved Edge PolicyManager provider path and require a nonempty value containing Army Bookmarks.'
        'Remote Apps Tasks|AVD-11' = 'Read HKLM:\SOFTWARE\Policies\Microsoft\Edge\NewTabPageLocation and require the Army SharePoint URL.'
        'Remote Apps Tasks|AVD-12' = 'Read HKLM/HKCU:\SOFTWARE\Policies\Microsoft\Edge\InPrivateModeAvailability and require DWORD 1.'
        'Remote Apps Tasks|AVD-14' = 'Read HKLM/HKCU:\SOFTWARE\Policies\Microsoft\Edge\ComposeInlineEnabled and require DWORD 0.'
        'Remote Apps Tasks|AVD-15' = 'Manual only: complete the IDCO CAC authentication workflow.'
        'Azure Virtual Desktop Tasks|AVD-09' = 'Manual only: launch DTP Lite and complete its required popup/user interaction.'
        'Azure Virtual Desktop Tasks|AVD-20' = 'Manual only: open About Teams and confirm AVD SlimCore Media Optimized or AVD Media Optimized is displayed.'
        'Azure Virtual Desktop Tasks|AVD-24' = 'Manual only: complete the documented Microsoft Teams activities and confirm Teams functions correctly.'
        'Azure Virtual Desktop Tasks|AVD-40' = 'Manual only: confirm Outlook Cached Exchange Mode is enabled and Download email for the past is set to 3 days.'
        'Azure Virtual Desktop Tasks|AVD-54' = 'Locate SnippingTool.exe through the current-user WindowsApps aliases/packages or System32, launch it, wait up to 30 seconds, then close the process started by the check.'
        'Azure Virtual Desktop Tasks|AVD-57' = 'Manual only: complete the IDCO CAC authentication workflow.'
        'Azure Virtual Desktop Tasks|AVD-58' = 'Manual only: visually confirm the Army wallpaper is applied in the signed-in AVD session.'
        'Azure Virtual Desktop Tasks|AVD-59' = 'Read BitLocker ProtectionStatus for C: and D:. Pass only when protection is Off on both drives, regardless of encryption state or percentage.'
        'Azure Virtual Desktop Tasks|AVD-60' = 'Manual only: as the signed-in standard UAT user, not SYSTEM or an elevated administrator, try to create a folder in C:\ and confirm access is denied.'
        'Azure Virtual Desktop Tasks|AVD-61' = 'Manual only: as the signed-in standard UAT user, not SYSTEM or an elevated administrator, try to create a folder in D:\ and confirm access is denied.'
        'Azure Virtual Desktop Tasks|AVD-62' = 'Require P:\ to be mounted, then attempt to create a uniquely named temporary file in its root as the UAT user; pass only when Windows returns access denied.'
        'Azure Virtual Desktop Tasks|AVD-64' = 'Launch File Explorer, Calculator, Notepad, Snipping Tool, Photos, and Paint one at a time; allow up to 30 seconds per app and close only newly started processes.'
        'Azure Virtual Desktop Tasks|AVD-66' = 'Read installed language inventory and require LpCab for every configured language. LXP packages and language features are not required.'
        'Azure Virtual Desktop Tasks|AVD-68' = 'Manual only: confirm OneDrive is accessible and syncing in the signed-in AVD session.'
        'Azure Virtual Desktop Tasks|AVD-70' = 'Read the local CurrentBuild and UBR, retrieve the official Microsoft update history, and require the local build to be greater than the latest build from the previous calendar month on the same build branch.'
        'Azure Virtual Desktop Tasks|AVD-72' = 'Require SetDisableUXWUAccess=1, DeferFeatureUpdates=1, and DeferQualityUpdates=1 under HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate. Manually confirm the rendered Windows Update restrictions and organization-policy message.'
        'Azure Virtual Desktop Tasks|AVD-76' = 'Manual only: browse to https://atrrs.army.mil, approve the user prompt, and confirm the site loads correctly.'
        'Azure Virtual Desktop Tasks|AVD-78' = 'Manual only: browse to https://gears.army.mil, approve the user prompt, authenticate, and complete the required upload/download check.'
        'Azure Virtual Desktop Tasks|AVD-82' = 'Read HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters\DisabledComponents (require DWORD 255/0xFF) and require every ms_tcpip6 adapter binding Enabled value to be False.'
    }
    if ($descriptions.ContainsKey($key)) { return $descriptions[$key] }
    return $Review
}

function Test-CommandAvailable {
    param([string]$Name)
    return $null -ne (Get-Command -Name $Name -ErrorAction SilentlyContinue)
}

function Invoke-BoundedApplicationCheck {
    param(
        [string]$FilePath,
        [string]$ProcessName,
        [string[]]$ArgumentList = @(),
        [ValidateRange(1, 300)]
        [int]$TimeoutSeconds = 30,
        [switch]$AllowExistingProcess
    )

    $existingIds = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $startedProcess = $null
    $newProcesses = @()
    try {
        $startParameters = @{ FilePath = $FilePath; PassThru = $true; ErrorAction = 'Stop' }
        if ($ArgumentList.Count -gt 0) { $startParameters.ArgumentList = $ArgumentList }
        $startedProcess = Start-Process @startParameters
        $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
        do {
            Start-Sleep -Milliseconds 250
            $newProcesses = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Where-Object { $_.Id -notin $existingIds })
        } until ($newProcesses.Count -gt 0 -or (Get-Date) -ge $deadline -or ($null -ne $startedProcess -and $startedProcess.HasExited))

        if ($newProcesses.Count -eq 0 -and $null -ne $startedProcess -and -not $startedProcess.HasExited -and $startedProcess.ProcessName -eq $ProcessName) {
            $newProcesses = @($startedProcess)
        }
        return [pscustomobject]@{
            Started = $newProcesses.Count -gt 0 -or ($AllowExistingProcess -and $existingIds.Count -gt 0)
            ProcessIds = @($newProcesses | Select-Object -ExpandProperty Id)
            WaitSeconds = $TimeoutSeconds
        }
    }
    finally {
        $ownedIds = @($newProcesses | Select-Object -ExpandProperty Id)
        if ($null -ne $startedProcess -and $startedProcess.Id -notin $existingIds) { $ownedIds += $startedProcess.Id }
        foreach ($processId in @($ownedIds | Sort-Object -Unique)) {
            Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-DnsResult {
    param([string]$Name)
    try {
        $records = @(Resolve-DnsName -Name $Name -ErrorAction Stop)
        if ($records.Count -gt 0) { return New-Result 'Partial Pass' "DNS resolved $Name ($($records.Count) records); application connectivity is not validated." 'Automated DNS evidence collected; manual connection task required' }
        return New-Result 'Fail' "DNS returned no records for $Name."
    }
    catch { return New-Result 'Fail' "DNS did not resolve ${Name}: $($_.Exception.Message)" }
}

function Get-WebConnectivityResult {
    param(
        [string]$HostName,
        [string]$Url,
        [string]$DisplayName
    )
    try { $records = @(Resolve-DnsName -Name $HostName -ErrorAction Stop) }
    catch { return New-Result 'Fail' "DNS did not resolve ${HostName}: $($_.Exception.Message)" }
    if ($records.Count -eq 0) { return New-Result 'Fail' "DNS returned no records for $HostName." }

    try {
        $response = Invoke-WebRequest -Uri $Url -Method Get -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop
        return New-Result 'Pass' "DNS resolved $HostName ($($records.Count) records), and $DisplayName returned HTTP $([int]$response.StatusCode). Connectivity is proven; authentication was not tested."
    }
    catch {
        $responseProperty = $_.Exception.PSObject.Properties['Response']
        if ($null -ne $responseProperty -and $null -ne $responseProperty.Value) {
            $statusCode = [int]$responseProperty.Value.StatusCode
            return New-Result 'Pass' "DNS resolved $HostName ($($records.Count) records), and $DisplayName returned HTTP $statusCode. Connectivity is proven; authentication was not tested."
        }
        return New-Result 'Fail' "DNS resolved $HostName, but $DisplayName did not return an HTTP response: $($_.Exception.Message)"
    }
}

function Get-FmswebResult {
    $hostName = 'fmsweb.fms.army.mil'
    try { $records = @(Resolve-DnsName -Name $hostName -ErrorAction Stop) }
    catch { return New-Result 'Fail' "FMSWEB DNS lookup failed: $($_.Exception.Message)" }

    $resolvedNames = @(
        $records | Select-Object -ExpandProperty Name -ErrorAction SilentlyContinue
        $records | Select-Object -ExpandProperty NameHost -ErrorAction SilentlyContinue
    ) | Where-Object { $_ } | Select-Object -Unique
    $dmzNames = @($resolvedNames | Where-Object { $_ -match 'dmz' })
    if ($dmzNames.Count -gt 0) { return New-Result 'Fail' "FMSWEB resolved through a disallowed DMZ alias: $($dmzNames -join ', '). Full DNS name chain: $($resolvedNames -join ', ')." }

    $connectivity = Get-WebConnectivityResult $hostName 'https://fmsweb.fms.army.mil/' 'the FMSWEB URL'
    if ($connectivity.Result -eq 'Fail') { return $connectivity }
    return New-Result 'Partial Pass' "FMSWEB DNS resolved without a DMZ alias. DNS name chain: $($resolvedNames -join ', '). $($connectivity.Notes) Authenticated FMSWEB activities were not tested." 'Automated DNS and HTTPS checks passed; manual authenticated activity check required'
}

function Get-BastionRdpResult {
    if ($isSystem) { return New-Result 'Manual' 'An authenticated Azure Bastion RDP session cannot be tested from SYSTEM context.' 'Manual task: run with -RunInteractiveChecks in the UAT user context' }
    if (-not $RunInteractiveChecks) { return New-Result 'Manual' 'The approved check requires an authenticated Azure Bastion RDP session.' 'Run with -RunInteractiveChecks to perform the guided portal test' }

    try { Start-Process -FilePath 'https://portal.azure.us/' -ErrorAction Stop | Out-Null }
    catch { return New-Result 'Manual' "The Azure Government portal could not be opened: $($_.Exception.Message)" 'Manual Bastion RDP session test required' }

    $confirmation = (Read-Host 'In Azure Government portal, connect to the target VM through Bastion. If the default Bastion fails, select another and retry. Enter S for a successful RDP session, F if no available Bastion succeeds, or X if the test is incomplete').Trim().ToUpperInvariant()
    if ($confirmation -eq 'S') { return New-Result 'Pass' 'The operator confirmed a successful RDP session to the target VM through Azure Bastion.' 'Operator-confirmed guided Bastion test passed' }
    if ($confirmation -eq 'F') { return New-Result 'Fail' 'The operator reported that no available Azure Bastion produced a successful RDP session to the target VM.' 'Guided Bastion test failed' }
    return New-Result 'Manual' 'The Azure Bastion RDP session test was not completed or its result was unclear.' 'Manual Bastion RDP session test remains required'
}

function Get-FieryUpdaterResult {
    $runPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\RunOnce'
    )
    if (-not $isSystem) {
        $runPaths += 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
        $runPaths += 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce'
    }

    $autorunFindings = @()
    foreach ($runPath in $runPaths) {
        $item = Get-ItemProperty -Path $runPath -ErrorAction SilentlyContinue
        if ($null -eq $item) { continue }
        foreach ($property in $item.PSObject.Properties) {
            if ($property.Name -match '^PS') { continue }
            $command = [string]$property.Value
            if ($property.Name -match '(?i)fiery|fpdu|fdpu' -or $command -match '(?i)fiery|fpdu|fdpu') {
                $autorunFindings += "$runPath :: $($property.Name)=$command"
            }
        }
    }

    if ($autorunFindings.Count -gt 0) { return New-Result 'Fail' "Fiery updater autorun references remain: $($autorunFindings -join '; ')." }
    $scopeNote = if ($isSystem) { 'Machine Run and RunOnce keys are clear; the signed-in UAT user HKCU keys were not available from SYSTEM.' } else { 'Machine and current-user Run and RunOnce keys are clear.' }
    return New-Result 'Partial Pass' "$scopeNote No value name or command matched Fiery, FPDU, or FDPU. Manually log out and back in, then confirm that no Fiery printer driver updater popup appears." 'Automated autorun scan passed; manual post-logon popup confirmation required'
}

function Get-DodEcmaCertificateResult {
    if ($isSystem) { return New-Result 'Manual' 'Current-user certificate stores cannot be assessed for the UAT user from SYSTEM.' 'Manual task: verify certificates in the UAT user context' }

    $certificates = @(Get-ChildItem Cert:\CurrentUser\CA, Cert:\CurrentUser\Root -ErrorAction SilentlyContinue)
    $dodCertificates = @($certificates | Where-Object { $_.Issuer -match '(?i)DOD' })
    $ecmaCertificates = @($certificates | Where-Object { $_.Issuer -match '(?i)ECMA' })
    if ($dodCertificates.Count -eq 0 -or $ecmaCertificates.Count -eq 0) {
        return New-Result 'Fail' "Required certificate family is missing. DoD certificates found: $($dodCertificates.Count); ECMA certificates found: $($ecmaCertificates.Count)."
    }

    $now = Get-Date
    $invalidCertificates = @($dodCertificates + $ecmaCertificates | Where-Object { $_.NotBefore -gt $now -or $_.NotAfter -lt $now } | Sort-Object Thumbprint -Unique)
    if ($invalidCertificates.Count -gt 0) {
        $invalidEvidence = @($invalidCertificates | ForEach-Object { "$($_.Thumbprint) [$($_.NotBefore.ToString('s')) to $($_.NotAfter.ToString('s'))]" })
        return New-Result 'Fail' "Found $($invalidCertificates.Count) DoD/ECMA certificate(s) outside their validity window: $($invalidEvidence -join '; ')."
    }

    return New-Result 'Partial Pass' "Found $($dodCertificates.Count) DoD and $($ecmaCertificates.Count) ECMA certificate(s) in the CurrentUser Intermediate Certification Authorities and Trusted Root Certification Authorities stores; all are currently valid. An approved thumbprint list was not supplied." 'Automated family and validity checks passed; approved-list completeness remains manual'
}

function Get-RemoteAppHostResult {
    $computerName = [string]$env:COMPUTERNAME
    if ([string]::IsNullOrWhiteSpace($computerName)) { return New-Result 'Manual' 'The current session host name is unavailable.' 'Manual ring and slice verification required' }
    if ([string]::IsNullOrWhiteSpace($ExpectedRemoteAppHostPattern)) {
        return New-Result 'Manual' "Current session host: $computerName. No approved Remote Apps host-name pattern was supplied." 'Supply -ExpectedRemoteAppHostPattern with the approved ring and slice naming regex'
    }

    try { $matchesExpectedHost = $computerName -match $ExpectedRemoteAppHostPattern }
    catch { return New-Result 'Manual' "The approved Remote Apps host-name regex is invalid: $($_.Exception.Message)" 'Correct -ExpectedRemoteAppHostPattern' }
    if ($matchesExpectedHost) { return New-Result 'Pass' "Remote Apps session host '$computerName' matches the approved ring and slice pattern '$ExpectedRemoteAppHostPattern'." }
    return New-Result 'Fail' "Remote Apps session host '$computerName' does not match the approved ring and slice pattern '$ExpectedRemoteAppHostPattern'."
}

function Get-DriveWriteResult {
    param([string]$Path, [string]$Id)
    if ($isSystem) { return New-Result 'Manual' 'This standard-user permission check is not valid in SYSTEM context.' 'Manual task: run in the UAT user context' }
    if (-not (Test-Path -LiteralPath $Path)) { return New-Result 'Fail' "The drive required for the permission test is not available: $Path" }
    $testFile = Join-Path $Path "AVD-UAT-$Id-$timestamp.tmp"
    try {
        New-Item -ItemType File -Path $testFile -ErrorAction Stop | Out-Null
    }
    catch {
        if ($_.CategoryInfo.Category -eq [System.Management.Automation.ErrorCategory]::PermissionDenied -or $_.Exception -is [System.UnauthorizedAccessException]) {
            return New-Result 'Pass' "Write was denied in $Path."
        }
        return New-Result 'Fail' "The write test for $Path failed for a reason other than access denial: $($_.CategoryInfo.Category) - $($_.Exception.Message)"
    }

    $cleanupNote = ''
    try { Remove-Item -LiteralPath $testFile -Force -ErrorAction Stop }
    catch { $cleanupNote = " Cleanup also failed: $($_.Exception.Message)" }
    return New-Result 'Fail' "Write unexpectedly succeeded in $Path.$cleanupNote"
}

function Get-PerformanceResult {
    $samples = [Math]::Ceiling($PerformanceSampleSeconds / $PerformanceSampleIntervalSeconds)
    $cpuReadings = @()
    $memoryReadings = @()
    $diskReadings = @()

    try {
        for ($sample = 0; $sample -lt $samples; $sample++) {
            $processor = Get-CimInstance -ClassName Win32_PerfFormattedData_PerfOS_Processor -Filter "Name='_Total'" -ErrorAction Stop
            $operatingSystem = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
            $disk = Get-CimInstance -ClassName Win32_PerfFormattedData_PerfDisk_PhysicalDisk -Filter "Name='_Total'" -ErrorAction Stop

            $cpuReadings += [double]$processor.PercentProcessorTime
            $memoryReadings += 100 * (1 - ([double]$operatingSystem.FreePhysicalMemory / [double]$operatingSystem.TotalVisibleMemorySize))
            $diskReadings += 1000 * [double]$disk.AvgDisksecPerTransfer

            if ($sample -lt ($samples - 1)) { Start-Sleep -Seconds $PerformanceSampleIntervalSeconds }
        }
    }
    catch {
        return New-Result 'Manual' "Performance counters could not be collected: $($_.Exception.Message)" 'Manual task: review performance in Task Manager'
    }

    $averageCpu = [Math]::Round(($cpuReadings | Measure-Object -Average).Average, 1)
    $maximumMemory = [Math]::Round(($memoryReadings | Measure-Object -Maximum).Maximum, 1)
    $averageDiskLatency = [Math]::Round(($diskReadings | Measure-Object -Average).Average, 1)
    $notes = "Sampled for $PerformanceSampleSeconds seconds. Average CPU: $averageCpu% (limit $MaximumAverageCpuPercent%); maximum memory used: $maximumMemory% (limit $MaximumMemoryUsedPercent%); average disk latency: $averageDiskLatency ms (limit $MaximumDiskLatencyMilliseconds ms)."

    if ($averageCpu -le $MaximumAverageCpuPercent -and $maximumMemory -le $MaximumMemoryUsedPercent -and $averageDiskLatency -le $MaximumDiskLatencyMilliseconds) {
        return New-Result 'Pass' $notes
    }
    return New-Result 'Fail' $notes
}

function Get-ApplicationStartupResult {
    param(
        [string]$Path,
        [string]$ProcessName,
        [string]$DisplayName,
        [string]$ManualLimitation
    )
    if ($isSystem) { return New-Result 'Manual' "$DisplayName startup cannot be assessed from SYSTEM context." 'Manual task: run in the UAT user context' }
    if (-not $RunInteractiveChecks) { return New-Result 'Manual' "$DisplayName startup requires an interactive user session and is not launched by unattended automation." "Manual task: launch and validate $DisplayName in the UAT user session" }
    if (-not (Test-Path -LiteralPath $Path)) { return New-Result 'Fail' "$DisplayName executable was not found: $Path" }

    try {
        $launch = Invoke-BoundedApplicationCheck -FilePath $Path -ProcessName $ProcessName
        if (-not $launch.Started) { return New-Result 'Fail' "$DisplayName did not start within $($launch.WaitSeconds) seconds." }
        return New-Result 'Partial Pass' "$DisplayName started within $($launch.WaitSeconds) seconds and the process created by this check was closed. $ManualLimitation" 'Automated bounded startup passed; manual visual confirmation required'
    }
    catch { return New-Result 'Fail' "$DisplayName startup failed: $($_.Exception.Message)" }
}

function Get-SnippingToolPath {
    $candidates = @(
        (Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\SnippingTool.exe'),
        'C:\Windows\System32\SnippingTool.exe'
    )
    $packagedAliases = @(Get-ChildItem -Path (Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\Microsoft.ScreenSketch_*\SnippingTool.exe') -File -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)
    return @($candidates + $packagedAliases | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } | Select-Object -Unique -First 1)
}

function Get-SnippingToolResult {
    $snippingToolPath = Get-SnippingToolPath | Select-Object -First 1
    if (-not $snippingToolPath) { return New-Result 'Fail' 'Snipping Tool was not found in the current-user WindowsApps aliases/packages or System32.' }
    return Get-ApplicationStartupResult $snippingToolPath 'SnippingTool' 'Snipping Tool' 'Manually run both snip tests and confirm each opens in Snipping Tool without prompting for a default application; if prompted, verify the MS-SCREENCLIP association.'
}

function Get-DtpLiteResult {
    return New-Result 'Manual' 'DTP Lite requires user interaction before it can fully launch, so unattended automation cannot validate this check.' 'Manual task: launch and validate DTP Lite in the UAT user session'
}

function Get-FoundryOdbcResult {
    $registryPaths = @(
        'HKLM:\SOFTWARE\ODBC\ODBCINST.INI',
        'HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBCINST.INI'
    )
    $registrations = @(
        foreach ($registryPath in $registryPaths) {
            Get-ChildItem -Path $registryPath -ErrorAction SilentlyContinue |
                Where-Object { $_.PSChildName -match 'Foundry|Palantir' } |
                ForEach-Object {
                    $properties = Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction SilentlyContinue
                    [pscustomobject]@{
                        Name = $_.PSChildName
                        Driver = [Environment]::ExpandEnvironmentVariables([string]$properties.Driver)
                    }
                }
        }
    )

    if ($registrations.Count -eq 0) { return New-Result 'Fail' 'No Foundry or Palantir ODBC driver registration was found in either registry view.' }
    $valid = @($registrations | Where-Object { $_.Driver -and (Test-Path -LiteralPath $_.Driver -PathType Leaf) })
    if ($valid.Count -eq 0) {
        $details = $registrations | ForEach-Object { "$($_.Name): $($_.Driver)" }
        return New-Result 'Fail' "Found ODBC registration, but no referenced driver DLL exists. Registrations: $($details -join '; ')."
    }
    $details = $valid | ForEach-Object { "$($_.Name): $($_.Driver)" }
    return New-Result 'Pass' "Found registered Foundry ODBC driver with an existing DLL: $($details -join '; ')."
}

function Get-ChromePurviewResult {
    if ([string]::IsNullOrWhiteSpace($PurviewExtensionId)) {
        return New-Result 'Manual' 'The authoritative Microsoft Purview Chrome extension ID was not supplied.' 'Manual task: supply -PurviewExtensionId from the approved Chrome policy'
    }
    if ($PurviewExtensionId -notmatch '^[a-p]{32}$') { return New-Result 'Fail' "The supplied Chrome extension ID is invalid: $PurviewExtensionId" }

    $policyPaths = @(
        'HKLM:\SOFTWARE\Policies\Google\Chrome\ExtensionInstallForcelist',
        'HKCU:\SOFTWARE\Policies\Google\Chrome\ExtensionInstallForcelist'
    )
    $policyValues = @(
        foreach ($policyPath in $policyPaths) {
            $policy = Get-ItemProperty -Path $policyPath -ErrorAction SilentlyContinue
            if ($null -ne $policy) {
                $policy.PSObject.Properties |
                    Where-Object { $_.Name -notlike 'PS*' } |
                    Select-Object -ExpandProperty Value
            }
        }
    )
    $isForced = @($policyValues | Where-Object { $_ -match "^$([regex]::Escape($PurviewExtensionId))(;|$)" }).Count -gt 0

    $userDataRoot = Join-Path $env:LOCALAPPDATA 'Google\Chrome\User Data'
    $manifests = @(Get-ChildItem -Path (Join-Path $userDataRoot "*\Extensions\$PurviewExtensionId\*\manifest.json") -File -ErrorAction SilentlyContinue)
    if ($isForced) {
        $profileEvidence = if ($manifests.Count -gt 0) { " Found $($manifests.Count) installed profile manifest(s)." } else { ' No user-profile manifest was required for this policy check.' }
        return New-Result 'Pass' "Purview extension $PurviewExtensionId is present in the Chrome ExtensionInstallForcelist policy.$profileEvidence"
    }
    if ($manifests.Count -gt 0) { return New-Result 'Fail' "Purview extension $PurviewExtensionId is installed but is not present in Chrome force-install policy, so enabled enforcement is not proven." }
    return New-Result 'Fail' "Purview extension $PurviewExtensionId is neither force-installed by policy nor installed in a local Chrome profile."
}

function Get-EdgePurviewResult {
    $matchingServices = @(
        Get-Service -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -match 'Microsoft.*Purview|MIP|DLP' }
    )
    if ($matchingServices.Count -eq 0) {
        return New-Result 'Fail' 'No Windows service with a display name matching Microsoft Purview, MIP, or DLP was found.'
    }
    $details = $matchingServices | ForEach-Object { "$($_.DisplayName) [$($_.Name), $($_.Status)]" }
    return New-Result 'Pass' "Found Purview/MIP/DLP service evidence: $($details -join '; ')."
}

function Get-EdgeCollectionsResult {
    $edgePaths = @(@(
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe')
    ) | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) })
    if ($edgePaths.Count -eq 0) { return New-Result 'Fail' 'Microsoft Edge executable was not found.' }

    $productVersion = (Get-Item -LiteralPath $edgePaths[0]).VersionInfo.ProductVersion
    if ($productVersion -notmatch '^(\d+)\.') { return New-Result 'Fail' "Microsoft Edge version could not be parsed: $productVersion" }
    $majorVersion = [int]$Matches[1]
    if ($majorVersion -ge 153) { return New-Result 'Pass' "Microsoft Edge $productVersion is version 153 or later, where Microsoft removed the Collections feature." }

    $policyPaths = @(
        'HKLM:\SOFTWARE\Policies\Microsoft\Edge',
        'HKCU:\SOFTWARE\Policies\Microsoft\Edge'
    )
    foreach ($policyPath in $policyPaths) {
        try {
            $enabled = Get-ItemPropertyValue -Path $policyPath -Name 'EdgeCollectionsEnabled' -ErrorAction Stop
            if ($enabled -eq 0) { return New-Result 'Pass' "Microsoft Edge $productVersion has Collections disabled by $policyPath\EdgeCollectionsEnabled=0." }
            return New-Result 'Fail' "Microsoft Edge $productVersion has EdgeCollectionsEnabled=$enabled at $policyPath; expected 0."
        }
        catch { }
    }
    return New-Result 'Fail' "Microsoft Edge $productVersion supports Collections, but the EdgeCollectionsEnabled=0 managed policy was not found."
}

function Get-EdgeDeveloperToolsResult {
    $policyPaths = @(
        'HKLM:\SOFTWARE\Policies\Microsoft\Edge',
        'HKCU:\SOFTWARE\Policies\Microsoft\Edge'
    )
    foreach ($policyPath in $policyPaths) {
        try {
            $availability = Get-ItemPropertyValue -Path $policyPath -Name 'DeveloperToolsAvailability' -ErrorAction Stop
            if ($availability -eq 2) { return New-Result 'Pass' "Microsoft Edge Developer Tools are disabled by $policyPath\DeveloperToolsAvailability=2." }
            return New-Result 'Fail' "Microsoft Edge has DeveloperToolsAvailability=$availability at $policyPath; expected 2."
        }
        catch { }
    }
    return New-Result 'Fail' 'Microsoft Edge DeveloperToolsAvailability=2 managed policy was not found.'
}

function Get-EdgeDownloadLocationResult {
    $policyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Edge'
    try { $promptValue = Get-ItemPropertyValue -Path $policyPath -Name 'PromptForDownloadLocation' -ErrorAction Stop }
    catch { return New-Result 'Fail' "Edge PromptForDownloadLocation was not found at $policyPath." }
    if ($promptValue -ne 1) { return New-Result 'Fail' "Edge PromptForDownloadLocation is '$promptValue'; expected 1." }

    $expectedDirectory = 'C:\Users\${user_name}\OneDrive - US Army\Downloads'
    $mdmWinsRoot = 'HKLM:\SOFTWARE\Microsoft\MDMWins'
    $configuredDirectories = @()
    foreach ($identity in @(Get-ChildItem -Path $mdmWinsRoot -ErrorAction SilentlyContinue)) {
        $downloadDirectoryPath = Join-Path $identity.PSPath 'Software\Policies\Microsoft\Edge\DownloadDirectory'
        try {
            $downloadDirectory = [string](Get-ItemPropertyValue -Path $downloadDirectoryPath -Name 'Value' -ErrorAction Stop)
            if (-not [string]::IsNullOrWhiteSpace($downloadDirectory)) {
                $configuredDirectories += @{ Source = "$($identity.Name)\Software\Policies\Microsoft\Edge\DownloadDirectory\Value"; Value = $downloadDirectory.Trim().TrimEnd('\') }
            }
        }
        catch { }
    }

    $match = $configuredDirectories | Where-Object { $_.Value -ieq $expectedDirectory } | Select-Object -First 1
    if ($null -ne $match) {
        return New-Result 'Pass' "Edge PromptForDownloadLocation=1 at $policyPath and MDMWins DownloadDirectory Value='$expectedDirectory' at $($match.Source)."
    }
    if ($configuredDirectories.Count -eq 0) {
        return New-Result 'Fail' "Edge PromptForDownloadLocation=1, but no nonempty DownloadDirectory Value was found beneath any $mdmWinsRoot identity."
    }
    $actualValues = $configuredDirectories | ForEach-Object { "'$($_.Value)' at $($_.Source)" }
    return New-Result 'Fail' "Edge PromptForDownloadLocation=1, but MDMWins DownloadDirectory did not equal '$expectedDirectory'. Found: $($actualValues -join '; ')."
}

function Get-EdgeHomePageResult {
    $expectedUrl = 'https://armyeitaas.sharepoint-mil.us'
    $policyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Edge'
    try { $newTabPageLocation = [string](Get-ItemPropertyValue -Path $policyPath -Name 'NewTabPageLocation' -ErrorAction Stop) }
    catch { return New-Result 'Fail' "Edge NewTabPageLocation was not found at $policyPath." }

    if ($newTabPageLocation.IndexOf($expectedUrl, [StringComparison]::OrdinalIgnoreCase) -lt 0) {
        return New-Result 'Fail' "Edge NewTabPageLocation is '$newTabPageLocation'; expected it to contain '$expectedUrl'."
    }
    return New-Result 'Pass' "Edge NewTabPageLocation at $policyPath contains $expectedUrl."
}

function Get-EdgeInPrivateResult {
    foreach ($policyPath in @('HKLM:\SOFTWARE\Policies\Microsoft\Edge', 'HKCU:\SOFTWARE\Policies\Microsoft\Edge')) {
        try {
            $availability = Get-ItemPropertyValue -Path $policyPath -Name 'InPrivateModeAvailability' -ErrorAction Stop
            if ($availability -eq 1) { return New-Result 'Pass' "Microsoft Edge InPrivate browsing is disabled by $policyPath\InPrivateModeAvailability=1." }
            return New-Result 'Fail' "Microsoft Edge has InPrivateModeAvailability=$availability at $policyPath; expected 1."
        }
        catch { }
    }
    return New-Result 'Fail' 'Microsoft Edge InPrivateModeAvailability=1 mandatory policy was not found.'
}

function Get-EdgeWritingAssistanceResult {
    foreach ($policyPath in @('HKLM:\SOFTWARE\Policies\Microsoft\Edge', 'HKCU:\SOFTWARE\Policies\Microsoft\Edge')) {
        try {
            $enabled = Get-ItemPropertyValue -Path $policyPath -Name 'ComposeInlineEnabled' -ErrorAction Stop
            if ($enabled -eq 0) { return New-Result 'Pass' "Microsoft 365 Copilot Rewrite writing assistance is disabled by $policyPath\ComposeInlineEnabled=0." }
            return New-Result 'Fail' "Microsoft Edge has ComposeInlineEnabled=$enabled at $policyPath; expected 0."
        }
        catch { }
    }
    return New-Result 'Fail' 'Microsoft Edge ComposeInlineEnabled=0 mandatory policy was not found.'
}

function Get-EdgeCbiiResult {
    if ($isSystem) { return New-Result 'Manual' 'The authenticated Edge CBII workflow cannot be tested from SYSTEM context.' 'Manual task: run with -RunInteractiveChecks in the UAT user context' }
    if (-not $RunInteractiveChecks) { return New-Result 'Manual' 'The approved check requires an interactive Menlo sign-in and redirect transaction in managed Edge.' 'Run with -RunInteractiveChecks to perform the guided CBII test' }

    $edgeCandidates = @(
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe')
    )
    $edgePath = $edgeCandidates | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } | Select-Object -First 1
    if (-not $edgePath) { return New-Result 'Fail' 'Microsoft Edge was not found for the CBII test.' }

    try { Start-Process -FilePath $edgePath -ArgumentList 'https://youtube.com' -ErrorAction Stop | Out-Null }
    catch { return New-Result 'Fail' "Microsoft Edge failed to open YouTube for the CBII test: $($_.Exception.Message)" }

    $confirmation = (Read-Host 'In Edge, complete the Menlo sign-in without sharing credentials. Enter S if YouTube redirected to Menlo, sign-in succeeded, and YouTube then loaded; F if any step failed; or X if incomplete').Trim().ToUpperInvariant()
    if ($confirmation -eq 'S') { return New-Result 'Pass' 'The operator confirmed that managed Edge redirected YouTube through Menlo, completed sign-in, and returned successfully to YouTube.' 'Operator-confirmed guided CBII test passed' }
    if ($confirmation -eq 'F') { return New-Result 'Fail' 'The operator reported that the managed Edge Menlo redirect, sign-in, or return to YouTube failed.' 'Guided CBII test failed' }
    return New-Result 'Manual' 'Managed Edge opened YouTube, but the Menlo redirect and sign-in transaction was not completed or its result was unclear.' 'Manual CBII transaction remains required'
}

function Get-OwaEncryptionResult {
    if ($isSystem) { return New-Result 'Manual' 'The authenticated OWA encryption workflow cannot be tested from SYSTEM context.' 'Manual task: run with -OwaUrl and -RunInteractiveChecks in the UAT user context' }
    if ([string]::IsNullOrWhiteSpace($OwaUrl)) { return New-Result 'Manual' 'The approved OWA endpoint was not supplied.' 'Supply -OwaUrl with the approved HTTPS OWA endpoint' }
    $parsedOwaUrl = $null
    if (-not [Uri]::TryCreate($OwaUrl, [UriKind]::Absolute, [ref]$parsedOwaUrl) -or $parsedOwaUrl.Scheme -ne 'https') {
        return New-Result 'Fail' "The approved OWA endpoint must be an absolute HTTPS URL: $OwaUrl"
    }
    if (-not $RunInteractiveChecks) { return New-Result 'Manual' "The approved OWA endpoint is $OwaUrl, but encrypted send and receive require an interactive authenticated user." 'Run with -OwaUrl and -RunInteractiveChecks to perform the guided OWA encryption test' }

    $edgeCandidates = @(
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe')
    )
    $edgePath = $edgeCandidates | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } | Select-Object -First 1
    if (-not $edgePath) { return New-Result 'Fail' 'Microsoft Edge was not found for the OWA encryption test.' }

    try { Start-Process -FilePath $edgePath -ArgumentList $parsedOwaUrl.AbsoluteUri -ErrorAction Stop | Out-Null }
    catch { return New-Result 'Fail' "Microsoft Edge failed to open the approved OWA endpoint: $($_.Exception.Message)" }

    $confirmation = (Read-Host 'In OWA, send an encrypted test email to an approved recipient, then receive and decrypt an encrypted reply. Enter S if both directions succeed, F if either fails, or X if incomplete').Trim().ToUpperInvariant()
    if ($confirmation -eq 'S') { return New-Result 'Pass' 'The operator confirmed successful encrypted email send and encrypted reply receipt/decryption in OWA.' 'Operator-confirmed guided OWA encryption test passed' }
    if ($confirmation -eq 'F') { return New-Result 'Fail' 'The operator reported that encrypted email send or encrypted reply receipt/decryption failed in OWA.' 'Guided OWA encryption test failed' }
    return New-Result 'Manual' 'OWA was opened, but the encrypted send and receive/decrypt transaction was not completed or its result was unclear.' 'Manual OWA encryption transaction remains required'
}

function Get-OceShortcutResult {
    if ($isSystem) { return New-Result 'Manual' 'Current-user OCE shortcuts and startup cannot be assessed from SYSTEM context.' 'Manual task: run in the UAT user context' }
    if (-not $RunInteractiveChecks) { return New-Result 'Manual' 'OCE shortcut and startup validation requires an interactive user session and is not launched by unattended automation.' 'Manual task: validate the OCE shortcuts and launch OCE in the UAT user session' }

    $desktopRoots = @(
        [Environment]::GetFolderPath('Desktop'),
        [Environment]::GetFolderPath('CommonDesktopDirectory')
    ) | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Container) }
    $startMenuRoots = @(
        [Environment]::GetFolderPath('Programs'),
        [Environment]::GetFolderPath('CommonPrograms')
    ) | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Container) }
    $desktopEntries = @(Get-ChildItem -Path $desktopRoots -Filter '*.lnk' -File -Force -ErrorAction SilentlyContinue | Where-Object { $_.BaseName -match '(?i)\bOCE\b' })
    $startMenuEntries = @(Get-ChildItem -Path $startMenuRoots -Filter '*.lnk' -File -Recurse -Force -ErrorAction SilentlyContinue | Where-Object { $_.BaseName -match '(?i)\bOCE\b' })
    $issues = @()
    if ($desktopEntries.Count -eq 0) { $issues += 'no OCE shortcut was found on the user or public desktop' }
    if ($startMenuEntries.Count -eq 0) { $issues += 'no OCE shortcut was found in the user or common Start Menu' }
    if ($issues.Count -gt 0) { return New-Result 'Fail' "OCE shortcut validation failed: $($issues -join '; ')." }

    $shortcut = @($startMenuEntries + $desktopEntries)[0]
    try {
        $launch = Invoke-BoundedApplicationCheck -FilePath $shortcut.FullName -ProcessName 'oce'
        if (-not $launch.Started) { return New-Result 'Fail' "OCE did not start within $($launch.WaitSeconds) seconds after launching $($shortcut.FullName)." }
        return New-Result 'Partial Pass' "Found $($desktopEntries.Count) desktop and $($startMenuEntries.Count) Start Menu OCE shortcut(s); OCE started within $($launch.WaitSeconds) seconds and the process created by this check was closed. Confirm the application window is responsive and no blocking dialog is visible." 'Automated bounded startup passed; manual visual confirmation required'
    }
    catch { return New-Result 'Fail' "OCE shortcut launch failed: $($_.Exception.Message)" }
}

function Get-OutlookCachedModeResult {
    if ($isSystem) { return New-Result 'Manual' 'The UAT user Outlook Cached Exchange Mode policy cannot be assessed from SYSTEM context.' 'Manual task: run in the UAT user context' }

    $policyPath = 'HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\Outlook\Cached Mode'
    $policy = Get-ItemProperty -Path $policyPath -ErrorAction SilentlyContinue
    if ($null -eq $policy) { return New-Result 'Fail' "Classic Outlook Cached Exchange Mode policy was not found: $policyPath" }

    $enableProperty = $policy.PSObject.Properties['Enable']
    $monthsProperty = $policy.PSObject.Properties['SyncWindowSetting']
    $daysProperty = $policy.PSObject.Properties['SyncWindowSettingDays']
    $enableValue = if ($null -ne $enableProperty) { $enableProperty.Value } else { '<missing>' }
    $monthsValue = if ($null -ne $monthsProperty) { $monthsProperty.Value } else { '<missing>' }
    $daysValue = if ($null -ne $daysProperty) { $daysProperty.Value } else { '<missing>' }
    $issues = @()
    if ($null -eq $enableProperty -or [int]$enableProperty.Value -ne 1) { $issues += "Enable is '$enableValue'; expected 1" }
    if ($null -eq $monthsProperty -or [int]$monthsProperty.Value -ne 0) { $issues += "SyncWindowSetting is '$monthsValue'; expected 0" }
    if ($null -eq $daysProperty -or [int]$daysProperty.Value -ne 3) { $issues += "SyncWindowSettingDays is '$daysValue'; expected 3" }
    if ($issues.Count -gt 0) { return New-Result 'Fail' "Classic Outlook Cached Exchange Mode policy mismatch: $($issues -join '; ')." }
    return New-Result 'Pass' "Classic Outlook Cached Exchange Mode is enabled and Download email for the past is enforced at 3 days by $policyPath."
}

function Get-PowerAutomateUpdateResult {
    $policyPath = 'HKLM:\SOFTWARE\Microsoft\Power Automate Desktop'
    try { $disabled = Get-ItemPropertyValue -Path $policyPath -Name 'DisableOptionalUpdates' -ErrorAction Stop }
    catch { return New-Result 'Fail' "Power Automate Desktop DisableOptionalUpdates policy was not found at $policyPath." }
    if ($disabled -eq 1) { return New-Result 'Pass' "Power Automate Desktop manual updates and update notifications are disabled by $policyPath\DisableOptionalUpdates=1." }
    return New-Result 'Fail' "Power Automate Desktop has DisableOptionalUpdates=$disabled at $policyPath; expected 1."
}

function Get-PowerShellSevenResult {
    $command = Get-Command -Name 'pwsh.exe' -ErrorAction SilentlyContinue
    if ($null -eq $command) { return New-Result 'Fail' 'PowerShell 7 executable pwsh.exe was not found.' }

    $outputFile = Join-Path $env:TEMP "AVD-UAT-pwsh-$([Guid]::NewGuid().ToString('N')).out"
    $errorFile = Join-Path $env:TEMP "AVD-UAT-pwsh-$([Guid]::NewGuid().ToString('N')).err"
    try {
        $scriptText = '$PSVersionTable.PSEdition + ''|'' + $PSVersionTable.PSVersion.ToString()'
        $process = Start-Process -FilePath $command.Source -ArgumentList @('-NoLogo', '-NoProfile', '-NonInteractive', '-Command', $scriptText) -Wait -PassThru -RedirectStandardOutput $outputFile -RedirectStandardError $errorFile -ErrorAction Stop
        $rawOutput = Get-Content -LiteralPath $outputFile -Raw -ErrorAction SilentlyContinue
        $rawErrorOutput = Get-Content -LiteralPath $errorFile -Raw -ErrorAction SilentlyContinue
        $output = if ($null -eq $rawOutput) { '' } else { $rawOutput.Trim() }
        $errorOutput = if ($null -eq $rawErrorOutput) { '' } else { $rawErrorOutput.Trim() }
        if ($process.ExitCode -ne 0) { return New-Result 'Fail' "PowerShell 7 exited with code $($process.ExitCode): $errorOutput" }
        if ($output -notmatch '^Core\|(\d+)\.(\d+)(?:\.|$)') { return New-Result 'Fail' "PowerShell 7 returned unexpected version output: $output" }
        if ([int]$Matches[1] -lt 7) { return New-Result 'Fail' "pwsh.exe reported unsupported version $($Matches[1]).$($Matches[2]); expected major version 7 or later." }
        return New-Result 'Pass' "PowerShell $($output.Split('|')[1]) started noninteractively, executed a command, and exited successfully."
    }
    catch { return New-Result 'Fail' "PowerShell 7 functional startup failed: $($_.Exception.Message)" }
    finally { Remove-Item -LiteralPath $outputFile, $errorFile -Force -ErrorAction SilentlyContinue }
}

function Get-GraphIntunePolicyResult {
    $requiredVersion = [Version]'2.25.0'
    $requiredModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.DeviceManagement',
        'Microsoft.Graph.DeviceManagement.Administration',
        'Microsoft.Graph.Identity.DirectoryManagement',
        'Microsoft.Graph.Security',
        'Microsoft.Graph.Beta.DeviceManagement'
    )
    $missing = @($requiredModules | Where-Object {
        $moduleName = $_
        -not (Get-Module -ListAvailable -Name $moduleName | Where-Object { $_.Version -eq $requiredVersion } | Select-Object -First 1)
    })
    if ($missing.Count -gt 0) { return New-Result 'Fail' "Microsoft Graph modules missing required version 2.25.0: $($missing -join ', ')." }
    if (-not $RunInteractiveChecks) { return New-Result 'Partial Pass' 'All required Microsoft Graph modules are available at version 2.25.0; delegated USGovDoD authentication and Intune policy output were not requested.' 'Run with -RunInteractiveChecks in the UAT user PowerShell 7 session to complete the check' }
    if ($isSystem) { return New-Result 'Manual' 'Delegated Graph authentication and PIM activation cannot run from SYSTEM context.' 'Manual task: run with -RunInteractiveChecks in the UAT user context' }
    if ($PSVersionTable.PSEdition -ne 'Core' -or $PSVersionTable.PSVersion.Major -lt 7) { return New-Result 'Fail' "The guided Graph check requires PowerShell 7; current host is $($PSVersionTable.PSEdition) $($PSVersionTable.PSVersion)." }

    try {
        foreach ($moduleName in $requiredModules) { Import-Module -Name $moduleName -RequiredVersion $requiredVersion -Force -ErrorAction Stop }
        Start-Process -FilePath 'https://portal.azure.us' -ErrorAction Stop | Out-Null
        [void](Read-Host 'Activate PAG-AVD-Intune-Read-Only PIM in the browser, then press Enter')

        $tenantId = 'fae6d70f-954b-4811-92b6-0530d6f84c43'
        $context = Get-MgContext -ErrorAction SilentlyContinue
        if ($null -eq $context -or $context.TenantId -ne $tenantId -or $context.Environment -ne 'USGovDoD') {
            Connect-MgGraph -Environment 'USGovDoD' -TenantId $tenantId -NoWelcome -ErrorAction Stop | Out-Null
        }
        $uri = "https://graph.microsoft.us/beta/deviceManagement/configurationPolicies('bfd22ead-81ab-49a2-bed9-c8eaa57960a8')/settings"
        $response = Invoke-MgRestMethod -Uri $uri -Method GET -ErrorAction Stop
        $settings = @($response.value | Where-Object { $null -ne $_.settingInstance })
        if ($settings.Count -eq 0) { return New-Result 'Fail' 'The Intune configuration policy query succeeded but returned no settingInstance output.' }
        return New-Result 'Pass' "Imported all required Microsoft Graph modules at version 2.25.0 and retrieved $($settings.Count) setting instance(s) from the approved Intune policy in USGovDoD."
    }
    catch { return New-Result 'Fail' "Guided Microsoft Graph Intune policy check failed: $($_.Exception.Message)" }
}

function Get-ChromeUpdatePolicyResult {
    $chromeUpdateValueName = 'Update{8A69D345-D564-463C-AFF1-A69D9E530F96}'
    $policyPaths = @(
        'HKLM:\SOFTWARE\Policies\Google\Update',
        'HKLM:\SOFTWARE\WOW6432Node\Policies\Google\Update'
    )

    foreach ($policyPath in $policyPaths) {
        if (-not (Test-Path -Path $policyPath)) { continue }
        $policy = Get-ItemProperty -Path $policyPath -ErrorAction SilentlyContinue
        $chromeSpecific = $policy.PSObject.Properties[$chromeUpdateValueName]
        if ($null -ne $chromeSpecific) {
            if ([int]$chromeSpecific.Value -eq 0) { return New-Result 'Pass' "Chrome updates are disabled by $policyPath\$chromeUpdateValueName = 0." }
            return New-Result 'Fail' "Chrome-specific update policy is not disabled: $policyPath\$chromeUpdateValueName = $($chromeSpecific.Value)."
        }
        $default = $policy.PSObject.Properties['UpdateDefault']
        if ($null -ne $default) {
            if ([int]$default.Value -eq 0) { return New-Result 'Pass' "Chrome updates are disabled by $policyPath\UpdateDefault = 0." }
            return New-Result 'Fail' "Default Google Update policy is not disabled: $policyPath\UpdateDefault = $($default.Value)."
        }
    }
    return New-Result 'Fail' 'No Chrome-specific or default Google Update policy was found to disable updates.'
}

function Test-ArmyBookmarksRegistryOutput {
    param([string[]]$RegistryOutput)
    return $null -ne ($RegistryOutput | Where-Object { $_ -match 'Army\s+Bookmarks' } | Select-Object -First 1)
}

function Test-ArmyBookmarksRegistryKey {
    param(
        [string]$BrowserName,
        [string]$RegistryKey
    )

    $regExe = Join-Path $env:WINDIR 'System32\reg.exe'
    if (-not (Test-Path -LiteralPath $regExe -PathType Leaf)) {
        return New-Result 'Fail' "Registry query utility was not found: $regExe"
    }
    $output = @(& $regExe query $RegistryKey /s /reg:64 2>&1)
    if ($LASTEXITCODE -ne 0) { return New-Result 'Fail' "$BrowserName registry key was not found or could not be read: $RegistryKey" }
    if (Test-ArmyBookmarksRegistryOutput -RegistryOutput $output) {
        return New-Result 'Pass' "$BrowserName registry key contains 'Army Bookmarks': $RegistryKey"
    }
    return New-Result 'Fail' "$BrowserName registry key exists, but its recursive content does not contain the string 'Army Bookmarks': $RegistryKey"
}

function Get-ChromeManagedBookmarksResult {
    $locations = @()
    $machinePolicyPath = 'HKLM:\SOFTWARE\Policies\Google\Chrome'
    try {
        $value = [string](Get-ItemPropertyValue -Path $machinePolicyPath -Name 'ManagedBookmarks' -ErrorAction Stop)
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $locations += @{ Source = "$machinePolicyPath\ManagedBookmarks"; Value = $value }
        }
    }
    catch { }

    $providersPath = 'HKLM:\SOFTWARE\Microsoft\PolicyManager\providers'
    foreach ($provider in @(Get-ChildItem -Path $providersPath -ErrorAction SilentlyContinue)) {
        $providerPolicyPath = Join-Path $provider.PSPath 'default\Device\chrome~Policy~googlechrome'
        try {
            $value = [string](Get-ItemPropertyValue -Path $providerPolicyPath -Name 'ManagedBookmarks' -ErrorAction Stop)
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                $locations += @{ Source = "$($provider.Name)\default\Device\chrome~Policy~googlechrome\ManagedBookmarks"; Value = $value }
            }
        }
        catch { }
    }

    $match = $locations | Where-Object { $_.Value -match 'Army\s+Bookmarks' } | Select-Object -First 1
    if ($null -ne $match) {
        return New-Result 'Pass' "Chrome ManagedBookmarks is nonempty and contains 'Army Bookmarks' at $($match.Source)."
    }
    if ($locations.Count -eq 0) {
        return New-Result 'Fail' 'Chrome ManagedBookmarks was empty or missing from both HKLM:\SOFTWARE\Policies\Google\Chrome and every PolicyManager provider chrome~Policy~googlechrome path.'
    }
    $sources = $locations | ForEach-Object { $_.Source }
    return New-Result 'Fail' "Chrome ManagedBookmarks was nonempty at $($sources -join '; '), but none of the values contained 'Army Bookmarks'."
}

function Get-EdgeManagedFavoritesResult {
    $providersPath = 'HKLM:\SOFTWARE\Microsoft\PolicyManager\providers'
    $locations = @()
    foreach ($provider in @(Get-ChildItem -Path $providersPath -ErrorAction SilentlyContinue)) {
        $providerPolicyPath = Join-Path $provider.PSPath 'default\Device\microsoft_edge~Policy~microsoft_edge'
        try {
            $managedFavorites = [string](Get-ItemPropertyValue -Path $providerPolicyPath -Name 'ManagedFavorites' -ErrorAction Stop)
            if (-not [string]::IsNullOrWhiteSpace($managedFavorites)) {
                $locations += @{ Source = "$($provider.Name)\default\Device\microsoft_edge~Policy~microsoft_edge\ManagedFavorites"; Value = $managedFavorites }
            }
        }
        catch { }
    }

    $match = $locations | Where-Object { $_.Value -match 'Army\s+Bookmarks' } | Select-Object -First 1
    if ($null -ne $match) {
        return New-Result 'Pass' "Edge ManagedFavorites is nonempty and contains 'Army Bookmarks' at $($match.Source)."
    }
    if ($locations.Count -eq 0) {
        return New-Result 'Fail' "Edge ManagedFavorites was empty or missing from every PolicyManager provider microsoft_edge~Policy~microsoft_edge path beneath $providersPath."
    }
    $sources = $locations | ForEach-Object { $_.Source }
    return New-Result 'Fail' "Edge ManagedFavorites was nonempty at $($sources -join '; '), but none of the values contained 'Army Bookmarks'."
}

function Get-CopilotArtifactsResult {
    if ($isSystem) {
        return New-Result 'Manual' 'Copilot desktop, Start menu, and browser favorites are user-scoped and cannot be assessed or repaired from SYSTEM context.' 'Manual task: run the automated workflow in the UAT user context'
    }
    if ([string]::IsNullOrWhiteSpace($CopilotArtifactName)) { $CopilotArtifactName = 'Microsoft Copilot Chat' }

    $desktop = [Environment]::GetFolderPath('Desktop')
    $startMenu = [Environment]::GetFolderPath('Programs')
    if ([string]::IsNullOrWhiteSpace($desktop) -or [string]::IsNullOrWhiteSpace($startMenu)) {
        return New-Result 'Manual' 'The current user desktop or Start menu Programs folder could not be resolved.' 'Manual task: verify the Copilot shortcuts'
    }

    $shortcutNames = @($CopilotArtifactName, 'Microsoft Copilot', 'Copilot')
    $findShortcut = {
        param([string]$Directory)
        foreach ($name in $shortcutNames) {
            foreach ($extension in @('.url', '.lnk')) {
                $candidate = Join-Path $Directory "$name$extension"
                if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
            }
        }
        return $null
    }
    $desktopShortcut = & $findShortcut $desktop
    $startMenuShortcut = & $findShortcut $startMenu
    $sourceShortcut = @($desktopShortcut, $startMenuShortcut) | Where-Object { $_ } | Select-Object -First 1
    $repairs = @()

    if (-not $desktopShortcut -and $sourceShortcut) {
        $desktopShortcut = Join-Path $desktop ([IO.Path]::GetFileName($sourceShortcut))
        Copy-Item -LiteralPath $sourceShortcut -Destination $desktopShortcut -Force -ErrorAction Stop
        $repairs += 'desktop shortcut'
    }
    if (-not $startMenuShortcut -and $sourceShortcut) {
        $startMenuShortcut = Join-Path $startMenu ([IO.Path]::GetFileName($sourceShortcut))
        Copy-Item -LiteralPath $sourceShortcut -Destination $startMenuShortcut -Force -ErrorAction Stop
        $repairs += 'Start menu shortcut'
    }

    if (-not $desktopShortcut -and -not [string]::IsNullOrWhiteSpace($CopilotUrl)) {
        $desktopShortcut = Join-Path $desktop "$CopilotArtifactName.url"
        Set-Content -LiteralPath $desktopShortcut -Value @('[InternetShortcut]', "URL=$CopilotUrl") -Encoding ASCII
        $repairs += 'desktop shortcut'
    }
    if (-not $startMenuShortcut -and -not [string]::IsNullOrWhiteSpace($CopilotUrl)) {
        $startMenuShortcut = Join-Path $startMenu "$CopilotArtifactName.url"
        Set-Content -LiteralPath $startMenuShortcut -Value @('[InternetShortcut]', "URL=$CopilotUrl") -Encoding ASCII
        $repairs += 'Start menu shortcut'
    }

    $missing = @()
    if (-not $desktopShortcut -or -not (Test-Path -LiteralPath $desktopShortcut -PathType Leaf)) { $missing += 'desktop' }
    if (-not $startMenuShortcut -or -not (Test-Path -LiteralPath $startMenuShortcut -PathType Leaf)) { $missing += 'Start menu' }
    if ($missing.Count -gt 0) {
        return New-Result 'Manual' "Copilot shortcut is missing from the $($missing -join ' and '), and no existing shortcut or configured CopilotUrl was available to create it." 'Manual task: supply -CopilotUrl or install an approved Copilot shortcut'
    }

    $bookmarkFiles = @(
        Get-ChildItem -Path (Join-Path $env:LOCALAPPDATA 'Google\Chrome\User Data\*\Bookmarks') -File -ErrorAction SilentlyContinue
        Get-ChildItem -Path (Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\User Data\*\Bookmarks') -File -ErrorAction SilentlyContinue
    )
    $favoriteFound = $false
    foreach ($bookmarkFile in $bookmarkFiles) {
        $content = Get-Content -LiteralPath $bookmarkFile.FullName -Raw -ErrorAction SilentlyContinue
        if (-not $content) { continue }
        if ($CopilotUrl -and $content -match [regex]::Escape($CopilotUrl)) { $favoriteFound = $true; break }
        if ($content -match [regex]::Escape($CopilotArtifactName)) { $favoriteFound = $true; break }
    }
    if (-not $favoriteFound -and $CopilotUrl) {
        foreach ($policyPath in @('HKLM:\SOFTWARE\Policies\Google\Chrome', 'HKCU:\SOFTWARE\Policies\Google\Chrome', 'HKLM:\SOFTWARE\Policies\Microsoft\Edge', 'HKCU:\SOFTWARE\Policies\Microsoft\Edge')) {
            $policy = Get-ItemProperty -Path $policyPath -ErrorAction SilentlyContinue
            if ($null -ne $policy -and (($policy.ManagedBookmarks -match [regex]::Escape($CopilotUrl)) -or ($policy.ManagedFavorites -match [regex]::Escape($CopilotUrl)))) {
                $favoriteFound = $true
                break
            }
        }
    }

    $repairNote = if ($repairs.Count -gt 0) { " Repaired: $($repairs -join ', ')." } else { '' }
    if ($favoriteFound) {
        return New-Result 'Pass' "Copilot shortcuts are present on the desktop and Start menu, and a matching Chrome or Edge favorite was found.$repairNote"
    }
    return New-Result 'Partial Pass' "Copilot shortcuts are present on the desktop and Start menu, but a matching Chrome or Edge favorite was not found.$repairNote" 'Automated shortcut validation passed; browser favorite confirmation remains required'
}

function Get-OfficeStartupResult {
    if ($isSystem) { return New-Result 'Manual' 'Microsoft Office startup cannot be assessed from SYSTEM context.' 'Manual task: run in the UAT user context' }
    if (-not $RunInteractiveChecks) { return New-Result 'Manual' 'Microsoft Office startup validation requires an interactive user session and is not launched by unattended automation.' 'Manual task: launch and validate each Office application in the UAT user session' }
    $officePath = 'C:\Program Files\Microsoft Office\root\Office16'
    $applications = @(
        @{ Name = 'Word'; File = 'WINWORD.EXE'; Process = 'WINWORD' },
        @{ Name = 'Outlook'; File = 'OUTLOOK.EXE'; Process = 'OUTLOOK' },
        @{ Name = 'PowerPoint'; File = 'POWERPNT.EXE'; Process = 'POWERPNT' },
        @{ Name = 'Excel'; File = 'EXCEL.EXE'; Process = 'EXCEL' },
        @{ Name = 'Visio'; File = 'VISIO.EXE'; Process = 'VISIO' },
        @{ Name = 'Project'; File = 'WINPROJ.EXE'; Process = 'WINPROJ' },
        @{ Name = 'OneNote'; File = 'ONENOTE.EXE'; Process = 'ONENOTE' }
    )
    $failures = @()
    $launched = @()
    foreach ($application in $applications) {
        $path = Join-Path $officePath $application.File
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            $failures += "$($application.Name) executable is missing"
            continue
        }
        try {
            $launch = Invoke-BoundedApplicationCheck -FilePath $path -ProcessName $application.Process
            if ($launch.Started) { $launched += $application.Name }
            else { $failures += "$($application.Name) did not start within $($launch.WaitSeconds) seconds" }
        }
        catch { $failures += "$($application.Name) launch failed: $($_.Exception.Message)" }
    }
    if ($failures.Count -gt 0) { return New-Result 'Fail' "Office startup checks failed: $($failures -join '; ')." }
    return New-Result 'Partial Pass' "$($launched -join ', ') started within 10 seconds each and each process created by this check was closed; confirm no licensing, profile, or error dialogs are visible." 'Automated bounded startup passed; manual visual confirmation required'
}

function Get-OfficeUpdatePolicyResult {
    $policyPaths = @(
        'HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate',
        'HKCU:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate'
    )
    $policyValue = $null
    $policySource = $null
    foreach ($policyPath in $policyPaths) {
        try {
            $policyValue = Get-ItemPropertyValue -Path $policyPath -Name 'enableautomaticupdates' -ErrorAction Stop
            $policySource = $policyPath
            break
        }
        catch { }
    }
    $clickToRunState = $null
    try { $clickToRunState = Get-ItemPropertyValue -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -Name 'UpdatesEnabled' -ErrorAction Stop }
    catch { }
    $supportingEvidence = if ($null -eq $clickToRunState) { 'Click-to-Run UpdatesEnabled is not present.' } else { "Click-to-Run UpdatesEnabled=$clickToRunState." }

    if ($null -eq $policySource) { return New-Result 'Fail' "Managed Office enableautomaticupdates policy was not found. $supportingEvidence" }
    if ([int]$policyValue -eq 0) { return New-Result 'Pass' "Office automatic updates are disabled by $policySource\enableautomaticupdates = 0. $supportingEvidence" }
    return New-Result 'Fail' "Office automatic updates are not disabled: $policySource\enableautomaticupdates = $policyValue. $supportingEvidence"
}

function Get-TeamsOptimizationPrerequisiteResult {
    $issues = @()
    $isWvdEnvironment = $null
    try { $isWvdEnvironment = Get-ItemPropertyValue -Path 'HKLM:\SOFTWARE\Microsoft\Teams' -Name 'IsWVDEnvironment' -ErrorAction Stop }
    catch { }
    if ($isWvdEnvironment -ne 1) { $issues += "IsWVDEnvironment is '$isWvdEnvironment'; expected 1" }

    $teamsInstalled = $null -ne (Get-Command -Name 'ms-teams.exe' -ErrorAction SilentlyContinue)
    if (-not $teamsInstalled -and (Test-CommandAvailable 'Get-AppxPackage')) {
        $teamsInstalled = $null -ne (Get-AppxPackage -Name 'MSTeams' -ErrorAction SilentlyContinue | Select-Object -First 1)
    }
    if (-not $teamsInstalled) { $issues += 'Microsoft Teams installation was not found' }

    $redirector = Get-Service -Name 'RDWebRTCSvc' -ErrorAction SilentlyContinue
    if ($null -eq $redirector) { $issues += 'Remote Desktop WebRTC Redirector Service is not installed' }
    elseif ($redirector.Status -ne 'Running') { $issues += "Remote Desktop WebRTC Redirector Service is $($redirector.Status); expected Running" }

    if ($issues.Count -gt 0) { return New-Result 'Fail' "Teams media optimization prerequisites failed: $($issues -join '; ')." }
    return New-Result 'Partial Pass' 'IsWVDEnvironment=1, Microsoft Teams is installed, and the Remote Desktop WebRTC Redirector Service is running. Confirm About Teams shows AVD SlimCore Media Optimized or AVD Media Optimized.' 'Automated host prerequisites passed; manual Teams banner confirmation required'
}

function Get-TeamsStartupResult {
    if ($isSystem) { return New-Result 'Manual' 'Microsoft Teams startup cannot be assessed from SYSTEM context.' 'Manual task: run in the UAT user context' }
    if (-not $RunInteractiveChecks) { return New-Result 'Manual' 'Microsoft Teams startup requires an interactive user session and is not launched by unattended automation.' 'Manual task: launch and validate Microsoft Teams in the UAT user session' }

    $teamsCommand = Get-Command -Name 'ms-teams.exe' -ErrorAction SilentlyContinue
    $teamsPackage = $null
    if ($null -eq $teamsCommand -and (Test-CommandAvailable 'Get-AppxPackage')) {
        $teamsPackage = Get-AppxPackage -Name 'MSTeams' -ErrorAction SilentlyContinue | Select-Object -First 1
    }
    if ($null -eq $teamsCommand -and $null -eq $teamsPackage) { return New-Result 'Fail' 'Microsoft Teams installation was not found.' }

    try {
        $target = if ($null -ne $teamsCommand) { $teamsCommand.Source } else { 'shell:AppsFolder\MSTeams_8wekyb3d8bbwe!MSTeams' }
        $launch = Invoke-BoundedApplicationCheck -FilePath $target -ProcessName 'ms-teams'
        if (-not $launch.Started) { return New-Result 'Fail' "Microsoft Teams did not start within $($launch.WaitSeconds) seconds." }
        return New-Result 'Partial Pass' "Microsoft Teams started within $($launch.WaitSeconds) seconds and the process created by this check was closed. Confirm sign-in and required chat, calling, meeting, and presence workflows manually." 'Automated bounded startup passed; functional workflows require manual validation'
    }
    catch { return New-Result 'Fail' "Microsoft Teams startup failed: $($_.Exception.Message)" }
}

function Get-TeamsPopOutChatResult {
    if ($isSystem) { return New-Result 'Manual' 'An authenticated Teams pop-out chat cannot be tested from SYSTEM context.' 'Manual task: run with -RunInteractiveChecks in the UAT user context' }
    if (-not $RunInteractiveChecks) { return New-Result 'Manual' 'The approved check requires an authenticated Teams user to pop out an existing chat.' 'Run with -RunInteractiveChecks to perform the guided pop-out chat test' }

    $startupResult = Get-TeamsStartupResult
    if ($startupResult.Result -in @('Fail', 'Manual')) { return $startupResult }
    $confirmation = (Read-Host 'In Teams, open an existing chat and select Pop out chat. Enter S if a usable separate chat window opens, F if it does not, or X if the test is incomplete').Trim().ToUpperInvariant()
    if ($confirmation -eq 'S') { return New-Result 'Pass' 'The operator confirmed that an existing Teams chat opened and worked in a separate pop-out window.' 'Operator-confirmed guided Teams test passed' }
    if ($confirmation -eq 'F') { return New-Result 'Fail' 'The operator reported that Teams did not open a usable separate pop-out chat window.' 'Guided Teams test failed' }
    return New-Result 'Manual' 'Teams launched, but the pop-out chat transaction was not completed or its result was unclear.' 'Manual pop-out chat test remains required'
}

function Get-TeamsRollupResult {
    $teamsIds = @('AVD-20', 'AVD-21', 'AVD-22', 'AVD-23')
    $teamsResults = @($results | Where-Object { $_.Sheet -eq 'Azure Virtual Desktop Tasks' -and $_.ID -in $teamsIds })
    if ($teamsResults.Count -ne $teamsIds.Count) { return New-Result 'Manual' 'The preceding Teams checks were not all available for the AVD-24 roll-up.' 'Manual task: review AVD-20 through AVD-23' }

    $summary = ($teamsResults | ForEach-Object { "$($_.ID)=$($_.Result)" }) -join ', '
    if (@($teamsResults | Where-Object { $_.Result -eq 'Fail' }).Count -gt 0) { return New-Result 'Fail' "Teams roll-up failed because at least one preceding check failed: $summary." }
    if (@($teamsResults | Where-Object { $_.Result -eq 'Pass' }).Count -eq $teamsIds.Count) { return New-Result 'Pass' "All preceding Teams checks passed: $summary." }
    if (@($teamsResults | Where-Object { $_.Result -in @('Pass', 'Partial Pass') }).Count -gt 0) { return New-Result 'Partial Pass' "Teams roll-up has supporting evidence but unresolved manual tasks remain: $summary." 'Roll-up of AVD-20 through AVD-23; complete their named manual tasks' }
    return New-Result 'Manual' "All preceding Teams checks require manual work: $summary." 'Manual task: complete AVD-20 through AVD-23'
}

function Get-FirefoxProxyPolicyResult {
    $registryPaths = @(
        'HKLM:\SOFTWARE\Policies\Mozilla\Firefox\Proxy',
        'HKCU:\SOFTWARE\Policies\Mozilla\Firefox\Proxy'
    )
    foreach ($registryPath in $registryPaths) {
        try {
            $locked = Get-ItemPropertyValue -Path $registryPath -Name 'Locked' -ErrorAction Stop
            if ($locked -eq 1) { return New-Result 'Partial Pass' "Firefox Proxy policy is locked at $registryPath. Confirm the Connection Settings controls are disabled in Firefox." 'Automated policy passed; manual UI confirmation required' }
            return New-Result 'Fail' "Firefox Proxy policy at $registryPath has Locked=$locked; expected 1."
        }
        catch { }
    }

    $policyFiles = @(
        (Join-Path $env:ProgramFiles 'Mozilla Firefox\distribution\policies.json'),
        (Join-Path ${env:ProgramFiles(x86)} 'Mozilla Firefox\distribution\policies.json')
    ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }
    foreach ($policyFile in $policyFiles) {
        try {
            $document = Get-Content -LiteralPath $policyFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            $policiesProperty = $document.PSObject.Properties['policies']
            $proxyProperty = if ($null -ne $policiesProperty) { $policiesProperty.Value.PSObject.Properties['Proxy'] } else { $null }
            $lockedProperty = if ($null -ne $proxyProperty) { $proxyProperty.Value.PSObject.Properties['Locked'] } else { $null }
            if ($null -ne $lockedProperty -and $lockedProperty.Value -eq $true) { return New-Result 'Partial Pass' "Firefox Proxy policy is locked in $policyFile. Confirm the Connection Settings controls are disabled in Firefox." 'Automated policy passed; manual UI confirmation required' }
            if ($null -ne $proxyProperty) { return New-Result 'Fail' "Firefox Proxy policy in $policyFile is not locked." }
        }
        catch { return New-Result 'Fail' "Firefox policy file could not be evaluated: $policyFile. $($_.Exception.Message)" }
    }

    return New-Result 'Fail' 'No locked Firefox Proxy enterprise policy was found in the registry or Firefox distribution policies.json.'
}

function Get-FirefoxUpdatePolicyResult {
    $registryPaths = @(
        'HKLM:\SOFTWARE\Policies\Mozilla\Firefox',
        'HKCU:\SOFTWARE\Policies\Mozilla\Firefox'
    )
    foreach ($registryPath in $registryPaths) {
        try {
            $disabled = Get-ItemPropertyValue -Path $registryPath -Name 'DisableAppUpdate' -ErrorAction Stop
            if ($disabled -eq 1) { return New-Result 'Pass' "Firefox application updates are disabled by $registryPath\DisableAppUpdate=1." }
            return New-Result 'Fail' "Firefox policy at $registryPath has DisableAppUpdate=$disabled; expected 1."
        }
        catch { }
    }

    $policyFiles = @(
        (Join-Path $env:ProgramFiles 'Mozilla Firefox\distribution\policies.json'),
        (Join-Path ${env:ProgramFiles(x86)} 'Mozilla Firefox\distribution\policies.json')
    ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }
    foreach ($policyFile in $policyFiles) {
        try {
            $document = Get-Content -LiteralPath $policyFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            $policiesProperty = $document.PSObject.Properties['policies']
            $disabledProperty = if ($null -ne $policiesProperty) { $policiesProperty.Value.PSObject.Properties['DisableAppUpdate'] } else { $null }
            if ($null -ne $disabledProperty -and $disabledProperty.Value -eq $true) { return New-Result 'Pass' "Firefox application updates are disabled by $policyFile." }
            if ($null -ne $disabledProperty) { return New-Result 'Fail' "Firefox policy in $policyFile has DisableAppUpdate=$($disabledProperty.Value); expected true." }
        }
        catch { return New-Result 'Fail' "Firefox policy file could not be evaluated: $policyFile. $($_.Exception.Message)" }
    }

    return New-Result 'Fail' 'Firefox DisableAppUpdate policy was not found in the registry or Firefox distribution policies.json.'
}

function Get-YubiKeyDetectionResult {
    return New-Result 'Manual' 'YubiKey detection requires the operator to confirm the connected key in YubiKey Manager.' 'Manual task: open YubiKey Manager and confirm the connected YubiKey is displayed'
}

function Get-IdcoCacResult {
    if ($isSystem) { return New-Result 'Manual' 'IDCO CAC authentication and CurrentUser certificate evidence cannot be assessed from SYSTEM context.' 'Manual task: run in the UAT user context' }

    $issues = @()
    $smartCardService = Get-Service -Name 'SCardSvr' -ErrorAction SilentlyContinue
    if ($null -eq $smartCardService) { $issues += 'Smart Card service is not installed' }
    elseif ($smartCardService.Status -ne 'Running') { $issues += "Smart Card service is $($smartCardService.Status); expected Running" }

    try {
        $readers = @(Get-CimInstance -ClassName Win32_PnPEntity -Filter "PNPClass = 'SmartCardReader'" -ErrorAction Stop)
        $healthyReaders = @($readers | Where-Object { $_.Status -eq 'OK' })
        if ($healthyReaders.Count -eq 0) { $issues += 'no healthy smart-card reader was detected by Windows PnP' }
    }
    catch { $issues += "smart-card reader inventory failed: $($_.Exception.Message)" }

    $now = Get-Date
    $acceptedEkus = @('1.3.6.1.5.5.7.3.2', '1.3.6.1.4.1.311.20.2.2')
    $certificates = @(
        Get-ChildItem -Path 'Cert:\CurrentUser\My' -ErrorAction SilentlyContinue |
            Where-Object {
                $_.HasPrivateKey -and $_.NotBefore -le $now -and $_.NotAfter -gt $now -and
                @($_.EnhancedKeyUsageList | Where-Object { $_.ObjectId -in $acceptedEkus }).Count -gt 0
            }
    )
    if ($certificates.Count -eq 0) { $issues += 'no valid CurrentUser certificate with a private key and Client Authentication or Smart Card Logon EKU was found' }

    if (-not (Test-CommandAvailable 'Test-NetConnection')) {
        $issues += 'Test-NetConnection is unavailable for the IDCO TCP prerequisite check'
    }
    else {
        try {
            $tcpConnected = Test-NetConnection -ComputerName 'idco.dmdc.osd.mil' -Port 443 -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction Stop
            if (-not $tcpConnected) { $issues += 'TCP port 443 on idco.dmdc.osd.mil is unreachable' }
        }
        catch { $issues += "IDCO TCP prerequisite check failed: $($_.Exception.Message)" }
    }

    if ($issues.Count -gt 0) { return New-Result 'Fail' "IDCO CAC prerequisites failed: $($issues -join '; ')." }
    $evidence = "Smart Card service is running, a healthy reader is present, $($certificates.Count) valid user smart-card/client-auth certificate(s) were found, and idco.dmdc.osd.mil:443 is reachable."
    if (-not $RunInteractiveChecks) { return New-Result 'Partial Pass' $evidence 'Run with -RunInteractiveChecks in the UAT user session to complete Edge CAC authentication and confirm no blockers' }

    $edgeCandidates = @(
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe')
    )
    $edgePath = $edgeCandidates | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } | Select-Object -First 1
    if (-not $edgePath) { return New-Result 'Fail' "$evidence Microsoft Edge executable was not found." }

    try { Start-Process -FilePath $edgePath -ArgumentList 'https://idco.dmdc.osd.mil/idco' -ErrorAction Stop | Out-Null }
    catch { return New-Result 'Fail' "$evidence Edge failed to open IDCO: $($_.Exception.Message)" }

    $confirmation = (Read-Host 'Complete CAC authentication without sharing the PIN. Enter P if IDCO opens with no errors or blockers, or F if it does not').Trim().ToUpperInvariant()
    if ($confirmation -eq 'P') { return New-Result 'Pass' "$evidence The operator confirmed successful IDCO CAC authentication in Edge with no errors or blocking actions." 'Automated prerequisites and operator-confirmed guided login passed' }
    if ($confirmation -eq 'F') { return New-Result 'Fail' "$evidence The operator reported an error or blocking action during IDCO CAC authentication in Edge." 'Guided login failed' }
    return New-Result 'Partial Pass' "$evidence Edge opened IDCO, but the operator did not record a pass or fail result." 'Manual CAC login confirmation remains required'
}

function Get-ArmyWallpaperResult {
    if ($isSystem) { return New-Result 'Manual' 'The UAT user wallpaper registry value cannot be assessed from SYSTEM context.' 'Manual task: run in the UAT user context' }

    $registryLocations = @(
        @{ Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\System'; Name = 'Wallpaper'; Type = 'policy' },
        @{ Path = 'HKCU:\Control Panel\Desktop'; Name = 'Wallpaper'; Type = 'desktop' }
    )
    $configuredWallpaper = $null
    $source = $null
    foreach ($location in $registryLocations) {
        try {
            $value = Get-ItemPropertyValue -Path $location.Path -Name $location.Name -ErrorAction Stop
            if (-not [string]::IsNullOrWhiteSpace([string]$value)) {
                $configuredWallpaper = [Environment]::ExpandEnvironmentVariables([string]$value)
                $source = "$($location.Type) registry value $($location.Path)\$($location.Name)"
                break
            }
        }
        catch { }
    }

    if (-not $configuredWallpaper) { return New-Result 'Fail' 'No wallpaper path is configured in the current user policy or desktop registry value.' }
    if ([System.IO.Path]::GetFileName($configuredWallpaper) -ine 'Army.png') {
        return New-Result 'Fail' "The $source references '$configuredWallpaper'; expected a file named Army.png."
    }
    if (-not (Test-Path -LiteralPath $configuredWallpaper -PathType Leaf)) {
        return New-Result 'Fail' "The $source references Army.png, but the file does not exist at '$configuredWallpaper'."
    }
    return New-Result 'Pass' "The $source references the existing Army wallpaper file '$configuredWallpaper'."
}

function Get-BitLockerProtectionResult {
    if (-not (Test-CommandAvailable 'Get-BitLockerVolume')) {
        return New-Result 'Manual' 'Get-BitLockerVolume is unavailable; BitLocker protection status could not be verified for C: and D:.' 'Manual task: confirm BitLocker protection is Off on both C: and D:'
    }

    $evidence = @()
    $protectedDrives = @()
    $unresolved = @()
    foreach ($drive in @('C:', 'D:')) {
        try {
            $volumes = @(Get-BitLockerVolume -MountPoint $drive -ErrorAction Stop)
            if ($volumes.Count -ne 1) {
                $unresolved += "${drive}: expected one BitLocker volume, received $($volumes.Count)"
                continue
            }
            $status = [string]$volumes[0].ProtectionStatus
            if ($status -in @('Off', '0')) {
                $evidence += "${drive}: BitLocker protection is Off"
            }
            elseif ($status -in @('On', '1')) {
                $evidence += "${drive}: BitLocker protection is On"
                $protectedDrives += $drive
            }
            else {
                $unresolved += "${drive}: BitLocker protection status is unknown ('$status')"
            }
        }
        catch { $unresolved += "${drive}: BitLocker protection status could not be read: $($_.Exception.Message)" }
    }

    $notes = (@($evidence) + @($unresolved)) -join '; '
    if ($protectedDrives.Count -gt 0) {
        return New-Result 'Fail' "$notes. Protection must be Off on both C: and D:; encryption state does not affect this requirement."
    }
    if ($unresolved.Count -gt 0) {
        return New-Result 'Manual' $notes 'Manual task: confirm BitLocker protection is Off on both C: and D:'
    }
    return New-Result 'Pass' "$notes. Both drives meet the BitLocker Off requirement, regardless of encryption state or percentage."
}

function Get-MicrosoftStoreBlockResult {
    if ($isSystem) { return New-Result 'Manual' 'Microsoft Store application control cannot be tested from SYSTEM context.' 'Manual task: run with -RunInteractiveChecks in the UAT user context' }
    if (-not $RunInteractiveChecks) { return New-Result 'Manual' 'The approved check requires a Microsoft Store launch attempt in the signed-in UAT user session.' 'Run with -RunInteractiveChecks to perform the guided application-control test' }

    try { Start-Process -FilePath 'ms-windows-store:' -ErrorAction Stop | Out-Null }
    catch { return New-Result 'Pass' "The Microsoft Store URI launch was blocked with an error: $($_.Exception.Message)" 'Guided launch test passed' }

    $confirmation = (Read-Host 'Enter B if Microsoft Store displayed the expected blocked response, or O if the Store opened and was usable').Trim().ToUpperInvariant()
    if ($confirmation -eq 'B') { return New-Result 'Pass' 'The operator confirmed that Microsoft Store displayed the expected blocked response.' 'Operator-confirmed guided launch test passed' }
    if ($confirmation -eq 'O') { return New-Result 'Fail' 'The operator confirmed that Microsoft Store opened and was usable instead of being blocked.' 'Guided launch test failed' }
    return New-Result 'Manual' 'Microsoft Store was invoked, but the operator did not record whether it was blocked or opened.' 'Manual application-control confirmation remains required'
}

function Get-WindowsAppsStartupResult {
    if ($isSystem) { return New-Result 'Manual' 'Windows application startup cannot be assessed from SYSTEM context.' 'Manual task: run in the UAT user context' }
    if (-not $RunInteractiveChecks) { return New-Result 'Manual' 'Windows application startup validation requires an interactive user session and is not launched by unattended automation.' 'Manual task: launch and validate the required Windows applications in the UAT user session' }

    $snippingToolPath = Get-SnippingToolPath | Select-Object -First 1
    $applications = @(
        @{ Name = 'File Explorer'; Target = 'explorer.exe'; Process = 'explorer'; AllowExisting = $true },
        @{ Name = 'Calculator'; Target = 'calc.exe'; Process = 'CalculatorApp' },
        @{ Name = 'Notepad'; Target = 'notepad.exe'; Process = 'notepad' },
        @{ Name = 'Snipping Tool'; Target = $snippingToolPath; Process = 'SnippingTool' },
        @{ Name = 'Photos'; Target = 'ms-photos:'; Process = 'Microsoft.Photos' },
        @{ Name = 'Paint'; Target = 'mspaint.exe'; Process = 'mspaint' }
    )
    $launched = @()
    $failures = @()
    foreach ($application in $applications) {
        if (-not $application.Target) {
            $failures += "$($application.Name): executable or application alias was not found"
            continue
        }
        try {
            $allowExisting = $application.ContainsKey('AllowExisting') -and $application.AllowExisting
            $launch = Invoke-BoundedApplicationCheck -FilePath $application.Target -ProcessName $application.Process -TimeoutSeconds 30 -AllowExistingProcess:$allowExisting
            if ($launch.Started) { $launched += $application.Name }
            else { $failures += "$($application.Name): did not start within $($launch.WaitSeconds) seconds" }
        }
        catch { $failures += "$($application.Name): $($_.Exception.Message)" }
    }

    if ($failures.Count -gt 0) {
        return New-Result 'Fail' "Windows application launch failed for $($failures -join '; '). Successfully invoked: $($launched -join ', ')."
    }
    return New-Result 'Partial Pass' "All required Windows applications started within 30 seconds and each process created by this check was closed: $($launched -join ', '). Manually confirm that each application is responsive and has no blocking error; dismiss any authentication request with its close button." 'Automated bounded launches passed; manual functional and visual confirmation required'
}

function Get-InstallerBlockResult {
    if ($isSystem) { return New-Result 'Manual' 'Installer application control cannot be tested from SYSTEM context.' 'Manual task: run with approved test files and -RunInteractiveChecks in the UAT user context' }
    if ([string]::IsNullOrWhiteSpace($ExeInstallerTestFile) -or [string]::IsNullOrWhiteSpace($MsiInstallerTestFile)) {
        return New-Result 'Manual' 'Approved benign EXE and MSI installer test files were not supplied.' 'Supply -ExeInstallerTestFile and -MsiInstallerTestFile'
    }

    $tests = @(
        @{ Type = 'EXE'; Path = $ExeInstallerTestFile; Extension = '.exe' },
        @{ Type = 'MSI'; Path = $MsiInstallerTestFile; Extension = '.msi' }
    )
    foreach ($test in $tests) {
        if ([System.IO.Path]::GetExtension($test.Path) -ine $test.Extension) { return New-Result 'Fail' "$($test.Type) test file does not have the expected $($test.Extension) extension: $($test.Path)" }
        if (-not (Test-Path -LiteralPath $test.Path -PathType Leaf)) { return New-Result 'Fail' "$($test.Type) test file was not found: $($test.Path)" }
    }
    if (-not $RunInteractiveChecks) { return New-Result 'Manual' 'The approved EXE and MSI test files exist, but runtime application-control attempts require an interactive UAT user session.' 'Run with -RunInteractiveChecks to perform both guided installer block tests' }

    $blocked = @()
    $allowed = @()
    $inconclusive = @()
    foreach ($test in $tests) {
        try { Unblock-File -LiteralPath $test.Path -ErrorAction Stop }
        catch { return New-Result 'Fail' "$($test.Type) test file could not be unblocked: $($_.Exception.Message)" }

        try { Start-Process -FilePath $test.Path -ErrorAction Stop | Out-Null }
        catch {
            $inconclusive += "$($test.Type) launch returned an error: $($_.Exception.Message)"
            continue
        }
        $confirmation = (Read-Host "Enter B if App Control blocked the $($test.Type) installer, A if the installer was allowed to start, or X if the result was unclear").Trim().ToUpperInvariant()
        if ($confirmation -eq 'B') { $blocked += $test.Type }
        elseif ($confirmation -eq 'A') { $allowed += $test.Type }
        else { $inconclusive += "$($test.Type) result was not confirmed" }
    }

    if ($allowed.Count -gt 0) { return New-Result 'Fail' "App Control allowed these unblocked installer types to start: $($allowed -join ', '). Blocked: $($blocked -join ', ')." 'Guided installer test failed' }
    if ($blocked.Count -eq $tests.Count) { return New-Result 'Pass' 'App Control blocked both approved, unblocked EXE and MSI installer test files.' 'Operator-confirmed guided installer tests passed' }
    return New-Result 'Manual' "Installer block testing was inconclusive: $($inconclusive -join '; '). Blocked: $($blocked -join ', ')." 'Manual application-control confirmation remains required'
}

function Get-ClipchampBlockResult {
    if ($isSystem) { return New-Result 'Manual' 'The Clipchamp Edge download and application-control workflow cannot run from SYSTEM context.' 'Manual task: run with -RunInteractiveChecks in the UAT user context' }
    if (-not $RunInteractiveChecks) {
        return New-Result 'Manual' 'The approved check requires downloading Microsoft Clipchamp through managed Edge and opening the downloaded installer in the signed-in UAT user session.' 'Run with -RunInteractiveChecks to perform the guided Clipchamp block test'
    }

    $edgeCandidates = @(
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe')
    )
    $edgePath = $edgeCandidates | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } | Select-Object -First 1
    if (-not $edgePath) { return New-Result 'Fail' 'Microsoft Edge was not found for the Clipchamp application-control test.' }

    try { Start-Process -FilePath $edgePath -ArgumentList 'https://apps.microsoft.com/' -ErrorAction Stop | Out-Null }
    catch { return New-Result 'Fail' "Edge failed to open Microsoft Apps: $($_.Exception.Message)" }

    $confirmation = (Read-Host 'In Edge, find Microsoft Clipchamp, download it, and open Microsoft Clipchamp Installer.exe. Enter B if App Control for Business blocked it, A if it was allowed, or X if the test was not completed').Trim().ToUpperInvariant()
    if ($confirmation -eq 'B') { return New-Result 'Pass' 'The operator confirmed that App Control for Business blocked Microsoft Clipchamp Installer.exe downloaded through Edge.' 'Operator-confirmed guided Clipchamp test passed' }
    if ($confirmation -eq 'A') { return New-Result 'Fail' 'The operator reported that Microsoft Clipchamp Installer.exe was allowed to run instead of being blocked.' 'Guided Clipchamp test failed' }
    return New-Result 'Manual' 'The Clipchamp download and application-control result was not confirmed.' 'Manual guided Clipchamp test remains required'
}

function Get-InstalledLanguageResult {
    if (-not (Test-CommandAvailable 'Get-InstalledLanguage')) {
        return New-Result 'Manual' 'Installed language inventory cannot be collected because Get-InstalledLanguage is unavailable.' 'Manual task: verify installed language packs'
    }

    try {
        $languageInventory = Get-InstalledLanguage -ErrorAction Stop
        $installedLanguages = @(foreach ($language in $languageInventory) { $language })
    }
    catch { return New-Result 'Manual' "Installed language inventory failed: $($_.Exception.Message)" 'Manual task: verify installed language packs' }
    $installedTags = @($installedLanguages | ForEach-Object {
        $tagProperty = $_.PSObject.Properties['LanguageTag']
        if ($null -eq $tagProperty) { $tagProperty = $_.PSObject.Properties['LanguageId'] }
        if ($null -ne $tagProperty) { [string]$tagProperty.Value }
    } | Where-Object { $_ } | Sort-Object -Unique)

    $requiredTags = @()
    if ($RequiredLanguages -and $RequiredLanguages.Count -gt 0) {
        $requiredTags = @($RequiredLanguages | ForEach-Object { ([string]$_.LanguageTag).Trim() })
        if (@($requiredTags | Where-Object { [string]::IsNullOrWhiteSpace($_) }).Count -gt 0) {
            return New-Result 'Fail' 'Language package validation failed: configuration entry has no LanguageTag.'
        }
    }
    elseif ($RequiredLanguageTags -and $RequiredLanguageTags.Count -gt 0) {
        $requiredTags = @($RequiredLanguageTags | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })
    }
    if ($requiredTags.Count -gt 0) {
        $requiredTags = @($requiredTags | Sort-Object -Unique)
        $problems = @()
        foreach ($tag in $requiredTags) {
            $installed = $installedLanguages | Where-Object {
                $tagProperty = $_.PSObject.Properties['LanguageTag']
                if ($null -eq $tagProperty) { $tagProperty = $_.PSObject.Properties['LanguageId'] }
                $null -ne $tagProperty -and [string]$tagProperty.Value -ieq $tag
            } | Select-Object -First 1
            if ($null -eq $installed) {
                $problems += "$tag is not installed"
                continue
            }

            $installedPacks = @($installed.LanguagePacks | ForEach-Object { ([string]$_).Split(',') } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            if ('LpCab' -notin $installedPacks) { $problems += "$tag is missing LpCab" }
        }
        if ($problems.Count -eq 0) {
            return New-Result 'Pass' "LpCab is installed for every required language: $($requiredTags -join ', '). LXP packages and language features are not required."
        }
        return New-Result 'Fail' "Language package validation failed: $($problems -join '; '). Installed language tags: $($installedTags -join ', ')."
    }

    return New-Result 'Manual' "Installed language tags: $($installedTags -join ', '). No usable approved required-language baseline was supplied." 'Supply -RequiredLanguages or -RequiredLanguageTags with every required BCP-47 language tag'
}

function Get-PuttyDownloadBlockResult {
    if ($isSystem) { return New-Result 'Manual' 'The PuTTY browser download test cannot run from SYSTEM context.' 'Manual task: run with -RunInteractiveChecks in the UAT user context' }
    if (-not $RunInteractiveChecks) { return New-Result 'Manual' 'The approved check requires a PuTTY download attempt in managed Edge.' 'Run with -RunInteractiveChecks to perform the guided download-block test' }

    $edgeCandidates = @(
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe')
    )
    $edgePath = $edgeCandidates | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } | Select-Object -First 1
    if (-not $edgePath) { return New-Result 'Fail' 'Microsoft Edge was not found for the PuTTY download test.' }

    try { Start-Process -FilePath $edgePath -ArgumentList 'https://putty.org/index.html' -ErrorAction Stop | Out-Null }
    catch { return New-Result 'Fail' "Edge failed to open the PuTTY site: $($_.Exception.Message)" }

    $confirmation = (Read-Host 'Attempt the approved PuTTY download in Edge without opening or running any file. Enter B if the download was blocked, A if it downloaded, or X if the result was unclear').Trim().ToUpperInvariant()
    if ($confirmation -eq 'B') { return New-Result 'Pass' 'The operator confirmed that managed Edge blocked the PuTTY application download.' 'Operator-confirmed guided download test passed' }
    if ($confirmation -eq 'A') { return New-Result 'Fail' 'The operator reported that managed Edge allowed the PuTTY application download.' 'Guided download test failed' }
    return New-Result 'Manual' 'The PuTTY download-block result was not confirmed.' 'Manual guided download test remains required'
}

function Get-OneDriveSyncPrerequisiteResult {
    if ($isSystem) { return New-Result 'Manual' 'The signed-in user OneDrive account and sync process cannot be assessed from SYSTEM context.' 'Manual task: run in the UAT user context' }

    $accountsPath = 'HKCU:\SOFTWARE\Microsoft\OneDrive\Accounts'
    $businessAccounts = @(
        Get-ChildItem -Path $accountsPath -ErrorAction SilentlyContinue |
            Where-Object { $_.PSChildName -like 'Business*' } |
            ForEach-Object {
                $account = Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction SilentlyContinue
                if ($null -ne $account -and -not [string]::IsNullOrWhiteSpace([string]$account.UserFolder)) { $account }
            }
    )
    if ($businessAccounts.Count -eq 0) { return New-Result 'Fail' 'No configured business OneDrive account with a UserFolder was found for the signed-in user.' }

    $accessibleRoots = @()
    $rootErrors = @()
    foreach ($account in $businessAccounts) {
        $root = [Environment]::ExpandEnvironmentVariables([string]$account.UserFolder)
        try {
            Get-Item -LiteralPath $root -ErrorAction Stop | Out-Null
            Get-ChildItem -LiteralPath $root -Force -ErrorAction Stop | Select-Object -First 1 | Out-Null
            $accessibleRoots += $root
        }
        catch { $rootErrors += "${root}: $($_.Exception.Message)" }
    }
    if ($accessibleRoots.Count -eq 0) { return New-Result 'Fail' "No configured business OneDrive root is accessible. $($rootErrors -join '; ')" }

    $processes = @(Get-Process -Name 'OneDrive' -ErrorAction SilentlyContinue)
    if ($processes.Count -eq 0) { return New-Result 'Fail' "Business OneDrive root is accessible at '$($accessibleRoots[0])', but OneDrive.exe is not running." }
    return New-Result 'Partial Pass' "Found $($businessAccounts.Count) configured business OneDrive account(s), $($accessibleRoots.Count) accessible sync root(s), and $($processes.Count) running OneDrive process(es). Cloud synchronization has not been proven." 'Local OneDrive prerequisites passed; manually confirm current cloud sync status'
}

function Get-OsBuildResult {
    $currentVersionPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    try { $currentVersion = Get-ItemProperty -Path $currentVersionPath -ErrorAction Stop }
    catch { return New-Result 'Fail' "Windows version metadata could not be read: $($_.Exception.Message)" }
    $build = [string]$currentVersion.CurrentBuild
    $revision = [string]$currentVersion.UBR
    $displayVersion = [string]$currentVersion.DisplayVersion
    $productName = [string]$currentVersion.ProductName
    if ([string]::IsNullOrWhiteSpace($build) -or [string]::IsNullOrWhiteSpace($revision)) { return New-Result 'Fail' 'Windows CurrentBuild or UBR metadata is missing.' }

    $localBuild = "$build.$revision"
    $localEvidence = "Current OS build: $productName $displayVersion, $localBuild."
    if ([int]$build -lt 22000) { return New-Result 'Fail' "$localEvidence Build $build is not a Windows 11 build branch." }

    if ([string]::IsNullOrWhiteSpace($WindowsUpdateHistoryUrl)) {
        if ([string]::IsNullOrWhiteSpace($ApprovedOsBuild)) {
            return New-Result 'Manual' "$localEvidence No Windows update-history URL or approved build was configured." 'Manual task: compare the displayed current build with the approved monthly release'
        }
        if ($ApprovedOsBuild -notmatch '^\d+\.\d+$') { return New-Result 'Fail' "ApprovedOsBuild must use full build.revision format; received '$ApprovedOsBuild'." }
        if ([version]$localBuild -ge [version]$ApprovedOsBuild) { return New-Result 'Pass' "$localEvidence Configured minimum build: $ApprovedOsBuild." }
        return New-Result 'Fail' "$localEvidence Configured minimum build: $ApprovedOsBuild."
    }

    try {
        $response = Invoke-WebRequest -Uri $WindowsUpdateHistoryUrl -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop
        $text = [Net.WebUtility]::HtmlDecode(($response.Content -replace '<[^>]+>', ' ' -replace '\s+', ' '))
        $previousMonth = (Get-Date).Date.AddMonths(-1)
        $monthName = $previousMonth.ToString('MMMM', [Globalization.CultureInfo]::InvariantCulture)
        $year = $previousMonth.Year
        $releasePattern = '(?i)(?<month>January|February|March|April|May|June|July|August|September|October|November|December)\s+(?<day>\d{1,2}),\s+(?<year>\d{4}).{0,400}?OS Builds?\s+(?<builds>[0-9.]+(?:\s+and\s+[0-9.]+)*)'
        $previousBuilds = @(
            [regex]::Matches($text, $releasePattern) |
                Where-Object { $_.Groups['month'].Value -eq $monthName -and [int]$_.Groups['year'].Value -eq $year } |
                ForEach-Object { [regex]::Matches($_.Groups['builds'].Value, '\d+\.\d+') | ForEach-Object { $_.Value } } |
                Where-Object { $_ -match "^$([regex]::Escape($build))\." } |
                ForEach-Object { [version]$_ } |
                Sort-Object -Unique
        )
        if ($previousBuilds.Count -eq 0) {
            return New-Result 'Manual' "$localEvidence Microsoft update history did not contain a $monthName $year build on branch $build." 'Manual task: compare the displayed current build with the Microsoft update history'
        }
        $previousBuild = $previousBuilds | Sort-Object -Descending | Select-Object -First 1
        $evidence = "$localEvidence Latest $monthName $year Microsoft build on branch ${build}: $previousBuild."
        if ([version]$localBuild -gt $previousBuild) { return New-Result 'Pass' $evidence }
        return New-Result 'Fail' "$evidence The current build must be greater than the previous month's latest build."
    }
    catch {
        return New-Result 'Manual' "$localEvidence Microsoft update history could not be retrieved or parsed: $($_.Exception.Message)" 'Manual task: use the displayed current build to decide Pass or Fail'
    }
}

function Get-WindowsUpdateUiPolicyResult {
    $policyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
    $expected = @{
        SetDisableUXWUAccess = 1
        DeferFeatureUpdates = 1
        DeferQualityUpdates = 1
    }
    $actual = @{}
    $mismatches = @()
    foreach ($name in $expected.Keys) {
        try { $actual[$name] = Get-ItemPropertyValue -Path $policyPath -Name $name -ErrorAction Stop }
        catch { $actual[$name] = "<unavailable: $($_.Exception.Message)>" }
        if ($actual[$name] -ne $expected[$name]) { $mismatches += "$name=$($actual[$name]) (expected 1)" }
    }

    if ($mismatches.Count -gt 0) { return New-Result 'Fail' "Windows Update UI restriction policies are missing or incorrect at ${policyPath}: $($mismatches -join '; ')." }
    return New-Result 'Partial Pass' "Required Windows Update policies are configured: SetDisableUXWUAccess=1, DeferFeatureUpdates=1, and DeferQualityUpdates=1 at $policyPath. Manually confirm that Check for updates and Pause updates are greyed out and the organization-policy message is visible." 'Required policies passed; manual rendered-UI confirmation required'
}

function Get-WindowsUpdateAdvancedOptionsResult {
    if ($isSystem) { return New-Result 'Manual' 'Windows Update Advanced options cannot be visually assessed from SYSTEM context.' 'Manual task: run with -RunInteractiveChecks in the UAT user context' }
    if (-not $RunInteractiveChecks) { return New-Result 'Manual' 'The approved check requires visual inspection of every displayed Windows Update Advanced option.' 'Run with -RunInteractiveChecks to perform the guided visual check' }

    try { Start-Process -FilePath 'ms-settings:windowsupdate-options' -ErrorAction Stop | Out-Null }
    catch { return New-Result 'Manual' "Windows Update Advanced options could not be opened: $($_.Exception.Message)" 'Manual visual check required' }

    $confirmation = (Read-Host 'Inspect every displayed Windows Update Advanced option. Enter G if all options are greyed out, E if any option is enabled, or X if the result is unclear').Trim().ToUpperInvariant()
    if ($confirmation -eq 'G') { return New-Result 'Pass' 'The operator confirmed that every displayed Windows Update Advanced option was greyed out.' 'Operator-confirmed guided visual check passed' }
    if ($confirmation -eq 'E') { return New-Result 'Fail' 'The operator reported that at least one displayed Windows Update Advanced option was enabled.' 'Guided visual check failed' }
    return New-Result 'Manual' 'Windows Update Advanced options were opened, but the operator did not confirm whether every displayed option was greyed out.' 'Manual visual confirmation remains required'
}

function Get-SevenZipResult {
    if ($isSystem) { return New-Result 'Manual' '7-Zip GUI startup cannot be assessed from SYSTEM context.' 'Manual task: run in the UAT user context' }
    if (-not $RunInteractiveChecks) { return New-Result 'Manual' '7-Zip GUI startup requires an interactive user session and is not launched by unattended automation.' 'Manual task: launch and validate 7-Zip File Manager in the UAT user session' }
    $sevenZip = Join-Path $env:ProgramFiles '7-Zip\7zFM.exe'
    if (-not (Test-Path -LiteralPath $sevenZip)) { return New-Result 'Fail' "7-Zip File Manager was not found: $sevenZip" }
    try {
        $launch = Invoke-BoundedApplicationCheck -FilePath $sevenZip -ProcessName '7zFM'
        if (-not $launch.Started) { return New-Result 'Fail' "7-Zip File Manager did not start within $($launch.WaitSeconds) seconds." }
        return New-Result 'Partial Pass' "7-Zip File Manager started within $($launch.WaitSeconds) seconds and the process created by this check was closed; visible error dialogs are not detectable by this process check." 'Automated bounded startup passed; manual visual confirmation required'
    }
    catch { return New-Result 'Fail' "7-Zip test failed: $($_.Exception.Message)" }
}

function Get-AutomatedResult {
    param(
        [string]$Sheet,
        [string]$Id
    )
    if ($Sheet -eq 'Remote Apps Tasks') {
        switch ($Id) {
            'AVD-01' { return Get-RemoteAppHostResult }
            'AVD-02' { return Get-TeamsStartupResult }
            'AVD-03' { return Get-TeamsPopOutChatResult }
            'AVD-04' { return New-Result 'Manual' 'Microsoft Teams microphone and speaker settings require inspection in the signed-in user interface.' 'Manual task: confirm Teams microphone and speaker selections' }
            'AVD-05' { return New-Result 'Manual' 'Microsoft Teams functionality requires an authenticated interactive user workflow.' 'Manual task: confirm Teams functions correctly' }
            'AVD-06' { return Get-EdgeCbiiResult }
            'AVD-07' { return Get-EdgeCollectionsResult }
            'AVD-08' { return Get-EdgeDeveloperToolsResult }
            'AVD-09' { return Get-EdgeDownloadLocationResult }
            'AVD-10' { return Get-EdgeManagedFavoritesResult }
            'AVD-11' { return Get-EdgeHomePageResult }
            'AVD-12' { return Get-EdgeInPrivateResult }
            'AVD-13' { return Get-OwaEncryptionResult }
            'AVD-14' { return Get-EdgeWritingAssistanceResult }
            'AVD-15' { return New-Result 'Manual' 'IDCO CAC validation requires an interactive browser session, CAC selection, PIN entry, and authenticated page confirmation.' 'Manual task: complete the documented IDCO CAC workflow' }
            'AVD-16' { return Get-ApplicationStartupResult 'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE' 'OUTLOOK' 'Classic Outlook' 'Confirm the profile connects and complete the required mailbox activities without errors.' }
            'AVD-17' { return Get-ClipchampBlockResult }
            default { return $null }
        }
    }
    if ($Sheet -ne 'Azure Virtual Desktop Tasks') { return $null }

    switch ($Id) {
        'AVD-00' { return Get-PerformanceResult }
        'AVD-01' { return New-Result 'Manual' 'Full validation requires an authenticated Microsoft Graph query for this managed device and every expected Intune policy.' 'Manual task: validate the device and expected policies in Intune or Microsoft Graph' }
        'AVD-02' { return Get-SevenZipResult }
        'AVD-03' { return Get-ApplicationStartupResult 'C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe' 'Acrobat' 'Adobe Acrobat' 'Confirm that no licensing prompt is visible.' }
        'AVD-04' { return Get-ApplicationStartupResult 'C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe' 'Acrobat' 'Adobe Acrobat' 'Complete the checklist activities and confirm they produce no errors.' }
        'AVD-05' { return Get-ApplicationStartupResult 'C:\Program Files\Amazon Web Services, Wickr\AWS WickrGov\AWSWickrGov.exe' 'AWSWickrGov' 'AWS WickrGov' 'Confirm that the sign-in prompt is visible.' }
        'AVD-06' { return New-Result 'Manual' 'Rendered banner text and background color require visual inspection in an interactive user session.' 'Manual task: verify the displayed banner text and background color' }
        'AVD-07' { return Get-ApplicationStartupResult 'C:\Program Files\DAX Studio\DaxStudio.exe' 'DaxStudio' 'DAX Studio' 'Confirm that no error dialog is visible.' }
        'AVD-08' {
            $expected = @('DEPTH_Bios', 'DEPTH_Cpu', 'DEPTH_ListeningProcess', 'DEPTH_LocalUser', 'DEPTH_LocalUserGroups', 'DEPTH_Memory', 'DEPTH_Network', 'DEPTH_OperatingSystem', 'DEPTH_Patches', 'DEPTH_PhysicalDisk', 'DEPTH_Proxy', 'DEPTH_RouteTable', 'DEPTH_Scaninfo', 'DEPTH_SecurityProducts', 'DEPTH_System', 'DEPTH_Tpminfo', 'DEPTH_Usb', 'DEPTH_WindowsOptionalFeatures')
            $uninstallPaths = @(
                'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
                'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
            )
            $installed = @(Get-ItemProperty -Path $uninstallPaths -ErrorAction SilentlyContinue)
            $found = @($expected | Where-Object {
                $expectedName = $_
                $installed | Where-Object {
                    $displayNameProperty = $_.PSObject.Properties['DisplayName']
                    $_.PSChildName -eq $expectedName -or ($null -ne $displayNameProperty -and $displayNameProperty.Value -eq $expectedName)
                } | Select-Object -First 1
            })
            $missing = @($expected | Where-Object { $_ -notin $found })
            if ($found.Count -ge 15) { return New-Result 'Pass' "Found $($found.Count) of $($expected.Count) expected DEPTH entries; the approved minimum is 15." }
            if ($found.Count -gt 0) { return New-Result 'Partial Pass' "Found $($found.Count) of $($expected.Count) expected DEPTH entries; the approved minimum is 15. Missing: $($missing -join ', ')." 'Automated evidence is below the approved threshold' }
            return New-Result 'Fail' "None of the $($expected.Count) expected DEPTH entries were found."
        }
        'AVD-09' { return Get-DtpLiteResult }
        'AVD-10' { return Get-FoundryOdbcResult }
        'AVD-11' { return New-Result 'Manual' 'The check requires an interactive Chrome session, Menlo redirect observation, user sign-in, and final YouTube URL verification.' 'Manual task: complete the documented Chrome redirect and sign-in workflow' }
        'AVD-12' { return Get-ChromePurviewResult }
        'AVD-13' { return Get-ChromeUpdatePolicyResult }
        'AVD-14' { return Get-ChromeManagedBookmarksResult }
        'AVD-15' {
            foreach ($policyPath in @('HKLM:\SOFTWARE\Policies\Google\Chrome', 'HKCU:\SOFTWARE\Policies\Google\Chrome')) {
                try {
                    $setting = Get-ItemPropertyValue -Path $policyPath -Name 'AIModeSettings' -ErrorAction Stop
                    if ([int]$setting -eq 1) { return New-Result 'Pass' "Chrome AI Mode is disabled by $policyPath\AIModeSettings=1." }
                    return New-Result 'Fail' "Chrome AIModeSettings is $setting at $policyPath; expected 1 (do not allow AI Mode integrations)."
                }
                catch { }
            }
            return New-Result 'Fail' 'Chrome AIModeSettings managed policy was not found; expected DWORD 1 under the machine or user Chrome policy key.'
        }
        'AVD-16' {
            $expected = @{ HistorySearchSettings = 2; CreateThemesSettings = 2; DevToolsGenAISettings = 2; AutoFillPredictionSettings = 2; GenAILocalFoundationalModelSettings = 1; HelpMeWriteSettings = 2; TabCompareSettings = 2 }
            $bad = @()
            $policyPaths = @('HKLM:\SOFTWARE\Policies\Google\Chrome', 'HKCU:\SOFTWARE\Policies\Google\Chrome')
            foreach ($name in $expected.Keys) {
                $actual = $null
                $source = $null
                foreach ($policyPath in $policyPaths) {
                    try {
                        $actual = Get-ItemPropertyValue -Path $policyPath -Name $name -ErrorAction Stop
                        $source = $policyPath
                        break
                    }
                    catch { }
                }
                if ($null -eq $source) { $bad += "$name is missing" }
                elseif ($actual -ne $expected[$name]) { $bad += "$name is $actual at $source; expected $($expected[$name])" }
            }
            if ($bad.Count -eq 0) { return New-Result 'Pass' 'All expected Chrome policy values match.' }
            return New-Result 'Fail' "Missing or mismatched Chrome policies: $($bad -join '; ')."
        }
        'AVD-17' { return Get-CopilotArtifactsResult }
        'AVD-18' { return Get-OfficeStartupResult }
        'AVD-19' { return Get-OfficeUpdatePolicyResult }
        'AVD-20' { return New-Result 'Manual' 'Teams media optimization requires inspection of About Teams in the signed-in AVD session.' 'Manual task: confirm About Teams displays AVD SlimCore Media Optimized or AVD Media Optimized' }
        'AVD-21' { return Get-TeamsStartupResult }
        'AVD-22' { return New-Result 'Manual' 'The check requires an authenticated Teams session, selection of an existing chat, Pop out chat, and visual confirmation that the separate chat window rendered correctly.' 'Manual task: complete the documented Teams pop-out chat workflow' }
        'AVD-23' { return New-Result 'Manual' 'The check requires an authenticated Teams session, inspection of microphone and speaker selectors, and confirmation of recording and playback.' 'Manual task: complete the documented Teams audio-device workflow' }
        'AVD-24' { return New-Result 'Manual' 'Microsoft Teams functionality requires an authenticated interactive user workflow; preceding check results do not prove this expectation.' 'Manual task: complete the documented Teams activities and confirm Teams functions correctly' }
        'AVD-25' { return New-Result 'Manual' 'The check requires an interactive Firefox session, Menlo redirect observation, operator CAC or sign-in, and final YouTube URL verification.' 'Manual task: complete the documented Firefox redirect and sign-in workflow' }
        'AVD-26' { return Get-FirefoxProxyPolicyResult }
        'AVD-27' { return Get-FirefoxUpdatePolicyResult }
        'AVD-28' { return Get-EdgePurviewResult }
        'AVD-29' { return New-Result 'Manual' 'The check requires two managed Edge cycles: observe the Menlo redirect, complete CAC authentication, verify the final YouTube URL, close and relaunch Edge, then repeat the challenge and redirect.' 'Manual task: complete both documented Edge redirect cycles' }
        'AVD-30' { return Get-EdgeDeveloperToolsResult }
        'AVD-31' { return Get-EdgeDownloadLocationResult }
        'AVD-32' { return Get-EdgeManagedFavoritesResult }
        'AVD-33' { return Get-EdgeHomePageResult }
        'AVD-34' { return Get-EdgeInPrivateResult }
        'AVD-35' { return New-Result 'Manual' 'The check requires authentication to OWA, a test recipient and subject, sending an encrypted message, receipt confirmation, and successful decryption.' 'Manual task: complete the documented OWA encrypted-message workflow' }
        'AVD-36' { return Get-EdgeWritingAssistanceResult }
        'AVD-37' { return Get-ApplicationStartupResult 'C:\Program Files\Notepad++\notepad++.exe' 'notepad++' 'Notepad++' 'Confirm that the editor window is responsive and no blocking dialog is visible.' }
        'AVD-38' { return Get-ApplicationStartupResult 'C:\ocetool\oce.exe' 'oce' 'OCE Tool' 'Confirm that the application window is responsive and no blocking dialog is visible.' }
        'AVD-39' { return Get-OceShortcutResult }
        'AVD-40' { return New-Result 'Manual' 'Outlook Cached Exchange Mode and the download period require inspection in the signed-in user profile.' 'Manual task: confirm Use Cached Exchange Mode to download email to an Outlook data file is checked and Download email for the past is set to 3 days' }
        'AVD-41' { return Get-ApplicationStartupResult 'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE' 'OUTLOOK' 'Classic Outlook' 'Confirm the profile connects and complete the required mailbox activities without errors.' }
        'AVD-42' { return Get-PowerAutomateUpdateResult }
        'AVD-43' { return Get-ApplicationStartupResult 'C:\Program Files\Microsoft Power BI Desktop\bin\PBIDesktop.exe' 'PBIDesktop' 'Power BI Desktop' 'Confirm that the report canvas is responsive and no sign-in, update, or error dialog is blocking the application.' }
        'AVD-44' { return Get-ApplicationStartupResult 'C:\Program Files\Power BI Report Builder\PowerBIReportBuilder.exe' 'PowerBIReportBuilder' 'Power BI Report Builder' 'Confirm that the report-design window is responsive and no blocking dialog is visible.' }
        'AVD-45' { return Get-PowerShellSevenResult }
        'AVD-46' { return Get-GraphIntunePolicyResult }
        'AVD-47' { return New-Result 'Manual' 'The check requires an approved test printer, pre-install queue inventory, PrinterLogic authentication, separate installation attempts in managed Edge and Chrome, and post-install queue and driver verification.' 'Manual task: complete the documented PrinterLogic workflow in both browsers' }
        'AVD-48' { return New-Result 'Manual' 'The check requires launch of Purebred, before-and-after CurrentUser certificate inventory, operator-controlled registration, and comparison against approved issuer, EKU, and subject criteria.' 'Manual task: complete Purebred registration and verify the resulting certificate' }
        'AVD-49' { return Get-ApplicationStartupResult 'C:\Program Files\SAP\FrontEnd\SAPGUI\SAPlogon.exe' 'saplogon' 'SAP Logon' 'Confirm that the SAP Logon window is responsive and no configuration or error dialog is blocking the application.' }
        'AVD-50' { return New-Result 'Manual' 'Open SelfHelp from the desktop and manually exercise each required action, confirming that every action completes successfully without an error or blocked UI.' 'Manual functional check required; the checklist does not define machine-verifiable actions or outcomes' }
        'AVD-51' { return New-Result 'Manual' 'The check requires creating a uniquely named test favorite or bookmark, running SelfHelp backup, removing the entry, running SelfHelp restore, and confirming the same entry returns.' 'Manual task: complete the documented SelfHelp browser backup and restore workflow' }
        'AVD-52' { return New-Result 'Manual' 'Use SelfHelp to enter OPSEC mode and confirm the classification banner closes and the Army logo is obscured. Exit OPSEC mode and confirm the banner and Army 4K logo return; if the wallpaper does not restore on the same host, repeat the final verification on the other host in the pool.' 'Guided visual transaction required; rendered desktop state cannot be proven from configuration alone' }
        'AVD-53' { return Get-SnippingToolResult }
        'AVD-54' { return New-Result 'Manual' 'Use Snipping Tool to capture a visible region containing varied colors or text, open the resulting snip, and confirm the captured image is visible rather than black or blank.' 'Manual visual capture check required; application startup does not prove captured pixel rendering' }
        'AVD-55' { return Get-ApplicationStartupResult 'C:\Program Files (x86)\STIG Viewer 3\STIG Viewer 3.exe' 'STIG Viewer 3' 'STIG Viewer' 'Confirm that the main window is responsive and no error dialog is visible.' }
        'AVD-56' { return Get-YubiKeyDetectionResult }
        'AVD-57' { return New-Result 'Manual' 'IDCO CAC validation requires an interactive browser session, CAC selection, PIN entry, and authenticated page confirmation.' 'Manual task: complete the documented IDCO CAC workflow' }
        'AVD-58' { return New-Result 'Manual' 'The rendered Army wallpaper requires visual inspection in the signed-in AVD session; configuration alone does not prove it is applied.' 'Manual task: visually confirm the Army wallpaper is applied' }
        'AVD-59' { return Get-BitLockerProtectionResult }
        'AVD-60' { return New-Result 'Manual' 'C:\ drive permissions must be tested as the signed-in standard UAT user; SYSTEM or elevated administrator write access does not prove standard-user restrictions.' 'Manual task: without elevation, try to create a folder in C:\ and confirm access is denied' }
        'AVD-61' { return New-Result 'Manual' 'D:\ drive permissions must be tested as the signed-in standard UAT user; SYSTEM or elevated administrator write access does not prove standard-user restrictions.' 'Manual task: without elevation, try to create a folder in D:\ and confirm access is denied' }
        'AVD-62' { return Get-MicrosoftStoreBlockResult }
        'AVD-63' { return Get-WindowsAppsStartupResult }
        'AVD-64' { return Get-InstallerBlockResult }
        'AVD-65' { return Get-ClipchampBlockResult }
        'AVD-66' { return Get-InstalledLanguageResult }
        'AVD-67' { return Get-PuttyDownloadBlockResult }
        'AVD-68' { return New-Result 'Manual' 'OneDrive accessibility and active syncing require verification in the signed-in AVD session.' 'Manual task: confirm OneDrive is accessible and syncing' }
        'AVD-69' { return New-Result 'Manual' 'Choose a OneDrive file, select Always keep on this device, and confirm its icon becomes a solid green check. Log out of the AVD session, wait about 30 seconds, log back in, and confirm the same file changes to a blue cloud within one minute. If status icons or the context-menu action are missing, open the file and repeat the logoff/logon transition before recording the result.' 'Guided two-session manual check required; the runner cannot span logoff and relogon' }
        'AVD-70' { return Get-OsBuildResult }
        'AVD-71' {
            if (-not (Test-CommandAvailable 'Get-LocalUser')) { return New-Result 'Manual' 'The Guest account cannot be inspected because Get-LocalUser is unavailable.' 'Manual task: verify the Guest account name' }
            $guestAccount = Get-LocalUser | Where-Object { $_.SID.Value -match '-501$' } | Select-Object -First 1
            if ($null -ne $guestAccount -and $guestAccount.Name -eq 'xGuest') { return New-Result 'Pass' 'The built-in Guest account (SID ending in -501) is named xGuest.' }
            $actualName = if ($null -eq $guestAccount) { 'not found' } else { $guestAccount.Name }
            return New-Result 'Fail' "The built-in Guest account is not named xGuest; actual name: $actualName."
        }
        'AVD-72' { return Get-WindowsUpdateUiPolicyResult }
        'AVD-73' { return Get-WindowsUpdateAdvancedOptionsResult }
        'AVD-74' { return Get-WebConnectivityResult 'prdweb.afmiscloud.army.mil' 'https://prdweb.afmiscloud.army.mil/AFMISWEBNG/AFMISWeb/AFMISWebForms/LoginForm.aspx' 'the AFMIS login URL' }
        'AVD-75' { return New-Result 'Manual' 'ATRRS requires an interactive browser prompt and user approval before the site can be validated.' 'Manual task: open https://atrrs.army.mil, approve the prompt, and validate the site' }
        'AVD-76' { return Get-FmswebResult }
        'AVD-77' { return New-Result 'Manual' 'GEARS requires an interactive browser prompt, user approval, authentication, and file transfer validation.' 'Manual task: open https://gears.army.mil, approve the prompt, and complete the upload/download check' }
        'AVD-78' { return Get-BastionRdpResult }
        'AVD-79' { return Get-FieryUpdaterResult }
        'AVD-80' { return Get-DodEcmaCertificateResult }
        'AVD-81' {
            $disabledComponentsPath = 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters'
            try { $disabledComponents = Get-ItemPropertyValue -Path $disabledComponentsPath -Name 'DisabledComponents' -ErrorAction Stop }
            catch { return New-Result 'Fail' "IPv6 registry setting was not found: $disabledComponentsPath\DisabledComponents; expected DWORD 255 (0xFF)." }
            try { $bindings = @(Get-NetAdapterBinding -ComponentID ms_tcpip6 -ErrorAction Stop) }
            catch { return New-Result 'Manual' "NetAdapter binding inventory cannot be collected: $($_.Exception.Message)" 'Manual task: verify IPv6 bindings' }
            if ($bindings.Count -eq 0) { return New-Result 'Manual' 'Get-NetAdapterBinding returned no IPv6 bindings, so the disabled state cannot be proven.' 'Manual task: verify IPv6 bindings' }
            $enabled = @($bindings | Where-Object { $_.Enabled })
            if ([int64]$disabledComponents -ne 255) { return New-Result 'Fail' "IPv6 DisabledComponents is $disabledComponents at $disabledComponentsPath; expected 255 (0xFF). Adapter binding Enabled values: $($bindings.Enabled -join ', ')." }
            if ($enabled.Count -eq 0) { return New-Result 'Pass' "IPv6 DisabledComponents=255 (0xFF), and IPv6 binding Enabled=False on all $($bindings.Count) returned adapters." }
            return New-Result 'Fail' "IPv6 DisabledComponents=255 (0xFF), but IPv6 binding Enabled=True on: $($enabled.Name -join ', ')."
        }
        default { return $null }
    }
}

if (-not (Test-Path -LiteralPath $InputFile)) { throw "Input checklist not found: $InputFile" }
if (-not (Test-Path -LiteralPath $WorkbookTool)) { throw "Workbook helper not found: $WorkbookTool" }
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null

$sourceCsv = Join-Path $OutputDirectory "AVD-UAT-Source-$timestamp.csv"
$resultsCsv = Join-Path $OutputDirectory "AVD Checklist Results - $timestamp.csv"
$outputWorkbook = Join-Path $OutputDirectory "AVD Checklist - Completed - $timestamp.xlsx"
$outputHtml = Join-Path $OutputDirectory "AVD UAT Report - $timestamp.html"

try {
    Write-Verbose "Exporting source checklist from '$InputFile' with '$WorkbookTool'."
    & $WorkbookTool export --input $InputFile --output $sourceCsv
    if ($LASTEXITCODE -ne 0) { throw "Workbook export failed with exit code $LASTEXITCODE." }
    $checks = @(Import-Csv -LiteralPath $sourceCsv)
    Write-Verbose "Loaded $($checks.Count) checklist rows."
    foreach ($check in $checks) {
        Write-Verbose "Starting check: Sheet='$($check.Sheet)' ID='$($check.ID)' Test='$($check.Test)'"
        try {
            $automated = Get-AutomatedResult -Sheet $check.Sheet -Id $check.ID
        }
        catch {
            Write-Verbose "Check exception: Sheet='$($check.Sheet)' ID='$($check.ID)' Error='$($_.Exception.Message)' Details='$($_ | Out-String)'"
            $automated = New-Result 'Fail' "Automated check raised an exception: $($_.Exception.Message)" 'Automated check error; review troubleshooting log'
        }
        if ($null -eq $automated) {
            $automated = New-Result 'Manual' 'No reliable automated action can prove this checklist expectation.' 'Manual task required'
        }
        Write-Verbose "Completed check: Sheet='$($check.Sheet)' ID='$($check.ID)' Result='$($automated.Result)' Notes='$($automated.Notes)'"
        $results += $check | Select-Object Sheet, ID, Category, Test,
            @{ Name = 'Expected'; Expression = { $_.'Expected Results' } },
            @{ Name = 'ValidationCheck'; Expression = { Get-ValidationCheckDescription $check.Sheet $check.ID $automated.Review } },
            @{ Name = 'Result'; Expression = { $automated.Result } },
            @{ Name = 'IssueReference'; Expression = { '' } },
            @{ Name = 'Notes'; Expression = { $automated.Notes } },
            @{ Name = 'AutomationReview'; Expression = { $automated.Review } }
    }

    Write-Verbose "Writing CSV evidence to '$resultsCsv'."
    $results | Export-Csv -LiteralPath $resultsCsv -NoTypeInformation -Encoding UTF8
    Write-Verbose "Updating workbook at '$outputWorkbook'."
    & $WorkbookTool update --input $InputFile --results $resultsCsv --output $outputWorkbook
    if ($LASTEXITCODE -ne 0) { throw "Workbook update failed with exit code $LASTEXITCODE." }

    $passing = @($results | Where-Object { $_.Result -eq 'Pass' })
    $partial = @($results | Where-Object { $_.Result -eq 'Partial Pass' })
    $manual = @($results | Where-Object { $_.Result -eq 'Manual' })
    $failed = @($results | Where-Object { $_.Result -eq 'Fail' })
    $columns = @('Sheet', 'ID', 'Category', 'Test', 'Expected', 'ValidationCheck', 'Result', 'Notes', 'AutomationReview')
    $passingTable = ($passing | Select-Object $columns | ConvertTo-Html -Fragment -PreContent '<h2>Successful automated checks</h2>') -replace '<table>', '<table class="status-table pass">'
    $partialTable = ($partial | Select-Object $columns | ConvertTo-Html -Fragment -PreContent '<h2>Partially completed checks</h2>') -replace '<table>', '<table class="status-table partial">'
    $manualTable = ($manual | Select-Object $columns | ConvertTo-Html -Fragment -PreContent '<h2>Manual task checks</h2>') -replace '<table>', '<table class="status-table manual">'
    $failedTable = ($failed | Select-Object $columns | ConvertTo-Html -Fragment -PreContent '<h2>Failed automated checks</h2>') -replace '<table>', '<table class="status-table fail">'
    $context = if ($isSystem) { 'SYSTEM (machine context)' } else { "$env:USERDOMAIN\$env:USERNAME" }
    $body = @"
<header><h1>AVD UAT Report</h1><p>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Computer: $env:COMPUTERNAME | Context: $context</p></header>
<section class="summary"><strong>Total: $($results.Count)</strong><span class="pass">Pass: $($passing.Count)</span><span class="partial">Partial Pass: $($partial.Count)</span><span class="fail">Fail: $($failed.Count)</span><span class="manual">Manual: $($manual.Count)</span></section>
<main>$($passingTable -join "`r`n")$($partialTable -join "`r`n")$($failedTable -join "`r`n")$($manualTable -join "`r`n")</main>
"@
    $head = @'
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<style>body{margin:0;background:#f4f6f8;color:#18212b;font:14px "Segoe UI",sans-serif}header{padding:24px 3%;background:#17324d;color:#fff}h1{margin:0 0 6px}.summary{display:flex;gap:12px;padding:16px 3%;background:#fff;border-bottom:1px solid #ccd5de}.summary span,.summary strong{padding:7px 10px}.pass{background:#d9f2df}.partial,.manual{background:#fff0c2}.fail{background:#f8d7da}main{padding:8px 3% 32px}h2{margin-top:28px}table{width:100%;border-collapse:collapse;background:#fff}th,td{padding:8px;border:1px solid #ccd5de;text-align:left;vertical-align:top}th{background:#e6edf3}.status-table.pass td{background:#d9f2df}.status-table.partial td,.status-table.manual td{background:#fff0c2}.status-table.fail td{background:#f8d7da}@media(max-width:800px){table{display:block;overflow-x:auto}.summary{flex-wrap:wrap}}</style>
'@
    Write-Verbose "Writing HTML report to '$outputHtml'."
    ConvertTo-Html -Title 'AVD UAT Report' -Head $head -Body $body | Set-Content -LiteralPath $outputHtml -Encoding UTF8
}
finally { Remove-Item -LiteralPath $sourceCsv -Force -ErrorAction SilentlyContinue }

Write-Output 'AVD UAT collection completed.'
Write-Output "HTML report:    $outputHtml"
Write-Output "Excel checklist: $outputWorkbook"
Write-Output "CSV evidence:   $resultsCsv"
Write-Output "Pass: $($passing.Count) | Partial Pass: $($partial.Count) | Fail: $($failed.Count) | Manual: $($manual.Count)"

# SIG # Begin signature block
# MIIHMAYJKoZIhvcNAQcCoIIHITCCBx0CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDkeW9mDh/ycjiM
# OFnD0kOZnZW40Uhs7sJ+819F2mcaIKCCBBwwggQYMIICgKADAgECAhARPahbg4yR
# u0x0RWe84V06MA0GCSqGSIb3DQEBCwUAMCQxIjAgBgNVBAMMGUFWRCBVQVQgVGVz
# dCBDb2RlIFNpZ25pbmcwHhcNMjYwOTIzMTk0MjUxWhcNMzEwOTIzMTk1MjUwWjAk
# MSIwIAYDVQQDDBlBVkQgVUFUIFRlc3QgQ29kZSBTaWduaW5nMIIBojANBgkqhkiG
# 9w0BAQEFAAOCAY8AMIIBigKCAYEAr3U0akZAjTYFpKBuGxG/XOZrFq1TiCKvagLt
# k5EbZnFvlNami9JqgzikK4A/wFJxQBdhaffseFIxS0T7HearbpAtDuQTVdkhcC03
# uBMs68EyzlggBzT2hOH7v9ebpQ3l0iz23ubNUXAtevgda1KVHfpNdTo1I1kSE6vt
# dWzVx4XyrJDC/2kK7yMpFZAw1B0VqrNOt2EQDrBAYofgIQxB6QEmp7b7O33qlIBq
# hGJFziEsTqrGbr0oCQKiXblwmUS1qKl+hMb93xkZ29WiC3PCOpnW+jwEN5zfGRqH
# 1Q0j1eVEoFTNWuEJ81N7HTi7tNAC71gO/4PNQJ0562UFkovR9fUX00UUBjKCBSqy
# nqYIKKqSErrOAFYepr4bOpheVPNsWxxwjxLWfBze92kX9IgNN3otMqTvVFY2nYXl
# hhbN217V3yGoqVmaA6a+AIiBYhXt3WegcDtP02MX/mZvbZpeHC6QpqlKHZVxgdth
# cN/5evLCaoGxxBO85VJwFM/JKbUpAgMBAAGjRjBEMA4GA1UdDwEB/wQEAwIHgDAT
# BgNVHSUEDDAKBggrBgEFBQcDAzAdBgNVHQ4EFgQUkggBwTJ4XHabRzNkW8BAJOxx
# UXAwDQYJKoZIhvcNAQELBQADggGBAAZBilHu50OctPP6mSsw/AxFk/gdNzSr/1dC
# Hi24+oU6t6bcZcRp1JJXKbipWLHfR1z3QMjuOzbwVEOlBfq0rQRygVVrKkH21MXs
# ubIjl1AWXdkwrnbJl0TpiCBDQBBP+obPS+eJY524ovOBY2tZvI3jgOxIVJDDMRwu
# hB/5BzR/n7V1VUFRzRPeVWTsgjWC+qbkbzrPMgqk1yOinBtFdNRiU67lpMR8OKZx
# UiMVhacc/u8BsEaqFVLJD2jcXe8tnjqYpFHUdb1J6pe0rooMf6oTbQtBtokM5hcx
# rh+kkvuQeY0V340WFGKw1DfTJE9HCnwkGPR94KgVauYWmF3dF3qDlvAT6mUqZxT2
# LyR98EnVSlbS+cFWf/2jSEGAZlbv1axCS1JHaypwF3Lo0JUGbzkgE4eupgv65w0L
# bzOkrUw/s3ZTVzuzfHW4KFiGOhzwqO26lsviuWZtefpsPB7mhp6B65wo7Bt+wTrx
# Nbarmp4AGEZFrhszz4hGkaX1swYgaDGCAmowggJmAgEBMDgwJDEiMCAGA1UEAwwZ
# QVZEIFVBVCBUZXN0IENvZGUgU2lnbmluZwIQET2oW4OMkbtMdEVnvOFdOjANBglg
# hkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3
# DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEV
# MC8GCSqGSIb3DQEJBDEiBCC0pSJRD3GnS/Eqrh/kwXh4yocJ458xVgAiRX5+agfF
# 4TANBgkqhkiG9w0BAQEFAASCAYBPjAK3aUlSdLgJB6oh/IFDnqmY0AAlvvitjQkr
# QTqd108N8w2Y03fdxA/RwmFE3r7oclNLyd4149UPOhk/UUFRDspImwLkMDFC86Fr
# BVuW37Qn1vk1dMPl7yVSfC0IigHDkQU2+vf+Zrk5JuADg5q38jrO31oRHKjLCWzR
# 2ZcP2TvblVQvIBd+K+cRAarvqa6NcG8W3992Pyox9txHqcs/mT+fhnKwNRO92T+6
# A7Gmi/NfvzRvR9gpoYzriBCDTBNcd8sevTjRwB+HtLtuOKgFwiSJk41fxxt2v6DA
# pQXxaZEfjB9B6hnH227q+jwj/cnhHCW02Qi4XO5akcOc+/83Q5w3C0/33IdZ6WGp
# jYZHHB66zu2+Vodp8+XZmfzgVED7lgO1CymdCv5A/HP0qFBKjCi6fH8lVpMan7Wl
# 3XnwHOQWtWtPTOn0RojuK9pwgAlaC01pvAqAiAejEteEGR75PO7QnZDq16SE+0OJ
# +WNvBDtQlKHL91M/d//Qzd4xsfo=
# SIG # End signature block
