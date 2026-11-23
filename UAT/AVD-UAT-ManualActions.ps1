function Add-ManualText {
    param([string]$Existing, [string]$Label, [string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $Existing }
    if ([string]::IsNullOrWhiteSpace($Existing)) { return "$Label$Value" }
    return "$Existing | $Label$Value"
}

function New-ManualAction {
    param(
        [string]$Name, [string]$Path, [string]$Instructions,
        [string[]]$Arguments = @(), [string[]]$AlternatePaths = @()
    )
    return @{
        Name = $Name; Path = $Path; Instructions = $Instructions
        Arguments = $Arguments; AlternatePaths = $AlternatePaths
    }
}

function New-ManualBrowserAction {
    param([ValidateSet('Chrome', 'Firefox', 'Edge')][string]$Browser, [string]$Url)
    $relative = switch ($Browser) {
        'Chrome' { 'Google\Chrome\Application\chrome.exe' }
        'Firefox' { 'Mozilla Firefox\firefox.exe' }
        'Edge' { 'Microsoft\Edge\Application\msedge.exe' }
    }
    New-ManualAction $Browser "C:\Program Files\$relative" `
        "In the managed $Browser profile, navigate to $Url for this check. Complete authentication yourself; never enter credentials or PINs in comments." `
        -Arguments @($Url) -AlternatePaths @("C:\Program Files (x86)\$relative")
}

function Get-ManualActions {
    param([string]$Sheet, [string]$Id)
    $office = 'C:\Program Files\Microsoft Office\root\Office16'
    if ($Sheet -eq 'Remote Apps Tasks') {
        switch ($Id) {
            { $_ -in @('AVD-02', 'AVD-03', 'AVD-04', 'AVD-05') } {
                return New-ManualAction 'Teams' -Instructions 'Open the published Teams Remote App yourself in the required test environment, not a local Teams instance.'
            }
            { $_ -in @('AVD-06', 'AVD-07', 'AVD-08', 'AVD-09', 'AVD-10', 'AVD-11', 'AVD-12', 'AVD-13', 'AVD-14', 'AVD-15', 'AVD-17') } {
                $url = switch ($Id) {
                    'AVD-06' { 'https://youtube.com' }
                    'AVD-09' { 'edge://settings/downloads' }
                    'AVD-10' { 'edge://favorites' }
                    'AVD-11' { 'https://armyeitaas.sharepoint-mil.us' }
                    'AVD-13' { 'https://webmail.apps.mil' }
                    'AVD-15' { 'https://idco.dmdc.osd.mil/idco' }
                    'AVD-17' { 'https://apps.microsoft.com' }
                    default { 'edge://policy' }
                }
                return New-ManualAction 'Edge' -Instructions "In the published Edge Remote App, navigate to $url and complete the documented check yourself. Do not substitute a local browser."
            }
            'AVD-16' { return New-ManualAction 'Outlook' -Instructions 'Open the published Outlook Remote App yourself and complete the mailbox activities there, not in local Outlook.' }
        }
        return
    }
    if ($Sheet -ne 'Azure Virtual Desktop Tasks') { return }
    switch ($Id) {
        'AVD-00' { return New-ManualAction 'Task Manager' "$env:WINDIR\System32\Taskmgr.exe" }
        'AVD-02' { return New-ManualAction '7-Zip' "$env:ProgramFiles\7-Zip\7zFM.exe" 'Perform the documented archive test using test files only; do not overwrite existing user files.' }
        { $_ -in @('AVD-03', 'AVD-04', 'AVD-05') } {
            return New-ManualAction 'Adobe Acrobat' 'C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe' 'Keep Acrobat open for related checks. Any save, close and reopen required by the PDF protection test is performed by you.'
        }
        'AVD-06' { return New-ManualAction 'AWS WickrGov' 'C:\Program Files\Amazon Web Services, Wickr\AWS WickrGov\AWSWickrGov.exe' }
        'AVD-08' { return New-ManualAction 'DAX Studio' 'C:\Program Files\DAX Studio\DaxStudio.exe' }
        'AVD-10' { return New-ManualAction 'DTP Lite' 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\DTP Lite\DTP Lite.lnk' }
        { $_ -in @('AVD-12', 'AVD-13', 'AVD-14', 'AVD-15', 'AVD-16', 'AVD-17') } {
            $url = switch ($Id) {
                'AVD-12' { 'https://youtube.com' }
                'AVD-13' { 'chrome://extensions' }
                'AVD-14' { 'chrome://settings/help' }
                'AVD-15' { 'chrome://bookmarks' }
                default { 'chrome://policy' }
            }
            return New-ManualBrowserAction 'Chrome' $url
        }
        'AVD-18' { return New-ManualAction 'Microsoft Copilot Chat' -Instructions 'Open the approved Copilot desktop/Start shortcut and verify the documented browser favorite yourself.' }
        { $_ -in @('AVD-19', 'AVD-20') } {
            foreach ($app in @(
                @('Word', 'WINWORD'), @('Outlook', 'OUTLOOK'), @('PowerPoint', 'POWERPNT'),
                @('Excel', 'EXCEL'), @('Visio', 'VISIO'), @('Project', 'WINPROJ'), @('OneNote', 'ONENOTE')
            )) {
                New-ManualAction $app[0] "$office\$($app[1]).EXE" 'Verify this Office application against this row, including licensing/update controls where required. Keep it open for its remaining checks.'
            }
            return
        }
        { $_ -in @('AVD-21', 'AVD-22', 'AVD-23', 'AVD-24', 'AVD-25') } {
            return New-ManualAction 'Teams' 'msteams:' 'Complete the documented Teams activities in the signed-in UAT session.'
        }
        { $_ -in @('AVD-26', 'AVD-27', 'AVD-28') } {
            $url = if ($Id -eq 'AVD-26') { 'https://youtube.com' } else { 'about:policies' }
            return New-ManualBrowserAction 'Firefox' $url
        }
        { $_ -in @('AVD-29', 'AVD-30', 'AVD-31', 'AVD-32', 'AVD-33', 'AVD-34', 'AVD-35', 'AVD-36', 'AVD-37') } {
            $url = switch ($Id) {
                'AVD-30' { 'https://youtube.com' }
                'AVD-32' { 'edge://settings/downloads' }
                'AVD-33' { 'edge://favorites' }
                'AVD-34' { 'https://armyeitaas.sharepoint-mil.us' }
                'AVD-36' { 'https://webmail.apps.mil' }
                default { 'edge://policy' }
            }
            return New-ManualBrowserAction 'Edge' $url
        }
        'AVD-38' { return New-ManualAction 'Notepad++' 'C:\Program Files\Notepad++\notepad++.exe' }
        { $_ -in @('AVD-39', 'AVD-40') } { return New-ManualAction 'OCE Tool' 'C:\ocetool\oce.exe' }
        { $_ -in @('AVD-41', 'AVD-42') } { return New-ManualAction 'Outlook' "$office\OUTLOOK.EXE" }
        'AVD-43' { return New-ManualAction 'Power Automate Desktop' 'C:\Program Files (x86)\Power Automate Desktop\dotnet\PAD.Console.Host.exe' }
        'AVD-44' { return New-ManualAction 'Power BI Desktop' 'C:\Program Files\Microsoft Power BI Desktop\bin\PBIDesktop.exe' }
        'AVD-45' { return New-ManualAction 'Power BI Report Builder' 'C:\Program Files\Power BI Report Builder\PowerBIReportBuilder.exe' }
        { $_ -in @('AVD-01', 'AVD-46', 'AVD-47') } {
            return New-ManualAction 'PowerShell 7' 'C:\Program Files\PowerShell\7\pwsh.exe' 'Perform only the documented PowerShell/Graph checks. Open the approved portal for PIM as needed. Module imports and authentication remain operator-controlled.'
        }
        'AVD-48' {
            New-ManualBrowserAction 'Edge' 'https://print.army.mil'
            New-ManualBrowserAction 'Chrome' 'https://print.army.mil'
            return
        }
        'AVD-49' { return New-ManualAction 'Purebred' -Instructions 'Launch Purebred yourself and complete the approved registration and before/after certificate comparison.' }
        'AVD-50' { return New-ManualAction 'SAP Logon' 'C:\Program Files\SAP\FrontEnd\SAPGUI\SAPlogon.exe' }
        { $_ -in @('AVD-51', 'AVD-52', 'AVD-53') } {
            return New-ManualAction 'SelfHelp' 'C:\Program Files\SelfHelp\AUDS_SelfHelp.exe' 'Open browsers as needed for backup/restore. Perform any required logoff or host switch yourself, using save and exit before leaving the session.'
        }
        { $_ -in @('AVD-54', 'AVD-55') } {
            return New-ManualAction 'Snipping Tool' "$env:WINDIR\System32\SnippingTool.exe"
        }
        'AVD-56' { return New-ManualAction 'STIG Viewer' 'C:\Program Files (x86)\STIG Viewer 3\STIG Viewer 3.exe' }
        'AVD-57' { return New-ManualAction 'YubiKey Manager' 'C:\Program Files\Yubico\YubiKey Manager\ykman-gui.exe' }
        'AVD-58' { return New-ManualBrowserAction 'Edge' 'https://idco.dmdc.osd.mil/idco' }
        { $_ -in @('AVD-61', 'AVD-62') } {
            return New-ManualAction 'File Explorer' "$env:WINDIR\explorer.exe" 'Test permissions without elevation, as the standard UAT user. A missing drive or another error does not prove access denied.' -Arguments @('shell:mycomputerfolder')
        }
        'AVD-63' { return New-ManualAction 'Microsoft Store' 'ms-windows-store:' }
        'AVD-64' {
            foreach ($app in @(
                @('File Explorer', "$env:WINDIR\explorer.exe"),
                @('Calculator', "$env:WINDIR\System32\calc.exe"),
                @('Notepad', "$env:WINDIR\System32\notepad.exe"),
                @('Snipping Tool', "$env:WINDIR\System32\SnippingTool.exe"),
                @('Photos', 'ms-photos:'), @('Paint', "$env:WINDIR\System32\mspaint.exe")
            )) {
                New-ManualAction $app[0] $app[1] 'Verify this application against the checklist expectation. Close only test windows, never the Explorer desktop shell.'
            }
            return
        }
        'AVD-65' { return New-ManualAction 'Installer block tests' -Instructions 'Run only the approved benign EXE/MSI test installers yourself and verify the documented block behavior.' }
        'AVD-66' { return New-ManualBrowserAction 'Edge' 'https://apps.microsoft.com' }
        'AVD-68' { return New-ManualBrowserAction 'Edge' 'https://putty.org/index.html' }
        { $_ -in @('AVD-69', 'AVD-70') } {
            return New-ManualAction 'OneDrive' 'C:\Program Files\Microsoft OneDrive\OneDrive.exe' `
                'Inspect sync and the documented file state. Never stop OneDrive sync. For a required logoff, save and exit first and resume after signing back in.' -Arguments @('/open')
        }
        'AVD-71' {
            New-ManualAction 'Windows version' "$env:WINDIR\System32\winver.exe"
            New-ManualBrowserAction 'Edge' 'https://learn.microsoft.com/en-us/windows/release-health/windows11-release-information'
            return
        }
        { $_ -in @('AVD-73', 'AVD-74') } {
            $uri = if ($Id -eq 'AVD-73') { 'ms-settings:windowsupdate' } else { 'ms-settings:windowsupdate-options' }
            return New-ManualAction 'Windows Update Settings' $uri "Navigate to $uri yourself if Settings is already open."
        }
        { $_ -in @('AVD-75', 'AVD-76', 'AVD-77', 'AVD-78') } {
            $url = switch ($Id) {
                'AVD-75' { 'https://prdweb.afmiscloud.army.mil' }
                'AVD-76' { 'https://www.atrrs.army.mil' }
                'AVD-77' { 'https://fmsweb.fms.army.mil' }
                'AVD-78' { 'https://gears.army.mil' }
            }
            return New-ManualBrowserAction 'Edge' $url
        }
        'AVD-79' { return New-ManualAction 'Bastion' -Instructions 'Complete the approved Bastion/RDP workflow yourself in the required test environment.' }
        'AVD-80' { return New-ManualAction 'Certificate Manager' "$env:WINDIR\System32\certmgr.msc" 'Inspect the documented user and machine certificate stores yourself; do not import or remove certificates.' }
    }
}

function Test-ManualReviewRequired {
    param($Check)
    return $Check.Result -in @('Partial Pass', 'Manual') -or
        @(Get-ManualActions $Check.Sheet $Check.ID).Count -gt 0
}

function Get-ManualReviewGroups {
    param([object[]]$Rows)
    $groups = @()
    $byKey = @{}
    foreach ($check in $Rows) {
        if (-not (Test-ManualReviewRequired $check)) { continue }
        $rowKey = "$($check.Sheet)|$($check.ID)"
        $actions = @(Get-ManualActions $check.Sheet $check.ID)
        if ($actions.Count -eq 0) { $actions = @(New-ManualAction "Observation - $($check.Test)") }
        for ($index = 0; $index -lt $actions.Count; $index++) {
            $action = $actions[$index]
            # Worksheets represent different test environments; never merge their sessions.
            $groupKey = "$($check.Sheet)|$($action.Name)"
            if (-not $byKey.ContainsKey($groupKey)) {
                $group = @{
                    Key = $groupKey; Sheet = $check.Sheet; Name = $action.Name
                    Action = $action; Tasks = @()
                }
                $byKey[$groupKey] = $group
                $groups += $group
            }
            $byKey[$groupKey].Tasks += @{
                Key = "$rowKey|$index"; RowKey = $rowKey; Check = $check
                Action = $action; StepCount = $actions.Count
            }
        }
    }
    return $groups
}

function Read-ManualChoice {
    param([string]$Prompt, [string[]]$Allowed)
    do {
        $choice = ([string](Read-Host $Prompt)).Trim().ToUpperInvariant()
        if ($choice -notin $Allowed) { Write-Warning "Choose one of: $($Allowed -join ', ')." }
    } until ($choice -in $Allowed)
    return $choice
}

function Read-ManualObservation {
    $choice = Read-ManualChoice 'Final result: [P]ass, [F]ail, [M]anual/unresolved, [S]kip, [Q]save and exit' @('P', 'F', 'M', 'S', 'Q')
    if ($choice -eq 'Q') { throw 'Manual review stopped by the operator. Resume using the displayed checkpoint path.' }
    if ($choice -eq 'S') { return @{ Result = 'Skip'; Notes = ''; IssueReference = '' } }
    do {
        $notes = ([string](Read-Host 'Manual evidence or comment (required; do not enter credentials or PINs)')).Trim()
        if ([string]::IsNullOrWhiteSpace($notes)) { Write-Warning 'A comment is required before this result can be recorded.' }
    } until (-not [string]::IsNullOrWhiteSpace($notes))
    $issue = if ($choice -eq 'F') { [string](Read-Host 'Issue reference (optional)') } else { '' }
    $result = switch ($choice) { 'P' { 'Pass' }; 'F' { 'Fail' }; default { 'Manual' } }
    return @{ Result = $result; Notes = $notes; IssueReference = $issue }
}

function Resolve-ManualWindowsAppPath {
    [CmdletBinding()]
    param([ValidateSet('Snipping Tool', 'Paint', 'Photos')][string]$Name, [string]$FallbackPath)
    $packageName = switch ($Name) {
        'Paint' { 'Microsoft.Paint' }
        'Photos' { 'Microsoft.Windows.Photos' }
        default { 'Microsoft.ScreenSketch' }
    }
    $executables = switch ($Name) {
        'Paint' { @('PaintApp\mspaint.exe', 'Paint\mspaint.exe', 'mspaint.exe') }
        'Photos' { @('Photos.exe') }
        default { @('SnippingTool\SnippingTool.exe', 'SnippingTool.exe') }
    }
    $programFiles = if ($env:ProgramW6432) { $env:ProgramW6432 } else { $env:ProgramFiles }
    $windowsApps = Join-Path $programFiles 'WindowsApps'
    try {
        $packages = @(Get-AppxPackage -Name $packageName -ErrorAction Stop |
            Where-Object { $_.InstallLocation -like "$windowsApps\*" } |
            Sort-Object { [version]$_.Version } -Descending)
        foreach ($package in $packages) {
            foreach ($executable in $executables) {
                $candidate = Join-Path $package.InstallLocation $executable
                if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
            }
        }
    }
    catch {
        Write-Warning "Could not query registered $Name packages: $($_.Exception.Message). Checking $windowsApps directly."
    }
    try {
        $directories = @(Get-ChildItem -LiteralPath $windowsApps -Filter "$($packageName)_*" -Directory -ErrorAction Stop |
            Where-Object { $_.Name -match ('^' + [regex]::Escape($packageName) + '_\d+\.\d+\.\d+\.\d+_') } |
            Sort-Object { [version]($_.Name.Split('_')[1]) } -Descending)
        foreach ($directory in $directories) {
            foreach ($executable in $executables) {
                $candidate = Join-Path $directory.FullName $executable
                if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
            }
        }
    }
    catch {
        Write-Warning "Could not inspect $windowsApps for ${Name}: $($_.Exception.Message). No packaged executable was resolved."
    }
    return $FallbackPath
}

function Start-ManualApplication {
    param($Action)
    if (-not $Action.Path) {
        Write-Host "Open the required test windows for $($Action.Name) yourself."
        return
    }
    $target = $Action.Path
    if ($Action.Name -in @('Snipping Tool', 'Paint', 'Photos')) {
        $target = Resolve-ManualWindowsAppPath -Name $Action.Name -FallbackPath $target
        if ($Action.Name -eq 'Photos' -and $target -eq $Action.Path) {
            throw 'Microsoft Windows Photos (Microsoft.Windows.Photos) executable not found in WindowsApps. OneDrive Photos and Photos Legacy are not substitutes.'
        }
    }
    $isUri = $target -match '^[a-z][a-z0-9+.-]*:' -and $target -notmatch '^[a-z]:'
    if (-not $isUri) {
        $paths = @($target) + @($Action.AlternatePaths)
        $found = @($paths | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1)
        if ($found.Count -eq 0) { throw "$($Action.Name) executable not found: $($paths -join ', ')" }
        $target = $found[0]
    }
    $parameters = @{ FilePath = $target; ErrorAction = 'Stop' }
    if ($Action.Arguments.Count -gt 0) { $parameters.ArgumentList = $Action.Arguments }
    Start-Process @parameters | Out-Null
    Write-Host 'Launch requested. Verify the visible application yourself; existing instances and background processes are never closed.'
}

function Confirm-ManualWindowClosure {
    param([string]$Name)
    Write-Host "All saved checks for $Name in this group are complete. Close only its test windows yourself. Preserve pre-existing work, the Explorer desktop shell, OneDrive sync and unrelated services." -ForegroundColor Yellow
    $choice = Read-ManualChoice '[C]onfirm test windows are closed and continue, or [Q]save and exit' @('C', 'Q')
    if ($choice -eq 'Q') { throw 'Window closure is unconfirmed. Progress was saved; close the test windows before resuming.' }
}

function Save-ManualCheckpoint {
    param($State, [string]$Path)
    $temporary = "$Path.$([guid]::NewGuid().ToString('N')).tmp"
    try {
        $State | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $temporary -Encoding UTF8 -ErrorAction Stop
        Invoke-ManualFileTool -Arguments @('replace', $temporary, $Path)
    }
    finally {
        if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -ErrorAction Stop }
    }
}

function Invoke-ManualFileTool {
    param([string[]]$Arguments)
    $fileTool = Join-Path $PSScriptRoot 'tools\AvdManualFileTool\AvdManualFileTool.exe'
    if (-not (Test-Path -LiteralPath $fileTool -PathType Leaf)) { throw "Manual file helper not found: $fileTool" }
    $output = @(& $fileTool @Arguments)
    if ($LASTEXITCODE -ne 0) { throw "Manual file helper failed with exit code $LASTEXITCODE." }
    return $output
}

function Get-ManualFileHash {
    param([string]$LiteralPath)
    $hash = @(Invoke-ManualFileTool -Arguments @('hash', $LiteralPath))
    if ($hash.Count -ne 1 -or $hash[0] -cnotmatch '^[0-9A-F]{64}$') { throw 'Manual file helper returned an invalid SHA-256 digest.' }
    return $hash[0]
}

function New-ManualCheckpoint {
    param([string]$InputFile, [string]$AutomatedResultsCsv, [object[]]$Rows)
    return @{
        Version = 3; InputFile = $InputFile; AutomatedResultsCsv = $AutomatedResultsCsv
        WorkbookHash = Get-ManualFileHash -LiteralPath $InputFile
        CsvHash = Get-ManualFileHash -LiteralPath $AutomatedResultsCsv
        Reviewer = "$env:USERDOMAIN\$env:USERNAME"; Computer = $env:COMPUTERNAME
        Rows = $Rows; Steps = @(); CompletedRows = @()
        ActiveGroup = ''; ActiveApplication = ''; ClosurePending = $false
        ReportTimestamp = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
        Outputs = @()
    }
}

function Read-ManualCheckpoint {
    param([string]$Path, [string]$InputFile, [string]$AutomatedResultsCsv, [object[]]$Rows)
    $state = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    $expected = New-ManualCheckpoint $InputFile $AutomatedResultsCsv $Rows
    foreach ($name in @('Version', 'InputFile', 'AutomatedResultsCsv', 'WorkbookHash', 'CsvHash', 'Reviewer', 'Computer')) {
        if ($state.$name -ne $expected.$name) { throw "Checkpoint $name does not match this workflow, user, machine or latest automated run. Use a new -CheckpointPath for a new review." }
    }
    if (@($state.Rows).Count -ne $Rows.Count) { throw 'Checkpoint row count does not match the automated evidence.' }
    for ($index = 0; $index -lt $Rows.Count; $index++) {
        foreach ($column in @('Sheet', 'ID', 'Category', 'Test', 'Expected')) {
            if ($state.Rows[$index].$column -ne $Rows[$index].$column) { throw "Checkpoint row $index does not match the automated evidence." }
        }
        if ($state.Rows[$index].Result -notin @('Pass', 'Partial Pass', 'Fail', 'Manual')) { throw "Checkpoint row $index contains an invalid result." }
    }
    $tasks = @((Get-ManualReviewGroups $Rows) | ForEach-Object { $_.Tasks })
    $seen = @{}
    foreach ($step in @($state.Steps)) {
        $matching = @($tasks | Where-Object Key -eq $step.Key)
        if ($seen.ContainsKey($step.Key) -or $matching.Count -ne 1 -or $matching[0].RowKey -ne $step.RowKey -or
            $matching[0].Action.Name -ne $step.Name -or $step.Result -notin @('Pass', 'Fail', 'Manual', 'Skip')) {
            throw 'Checkpoint contains an invalid or duplicate manual step.'
        }
        $seen[$step.Key] = $true
    }
    foreach ($rowKey in @($state.CompletedRows)) {
        $required = @($tasks | Where-Object RowKey -eq $rowKey)
        if ($required.Count -eq 0 -or @($required | Where-Object { -not $seen.ContainsKey($_.Key) }).Count -gt 0) {
            throw 'Checkpoint contains an incomplete completed row.'
        }
    }
    if ($state.ClosurePending) {
        Write-Warning "The interrupted application group was $($state.ActiveGroup). No saved PID is used. Confirm test-window closure before continuing."
        Confirm-ManualWindowClosure $state.ActiveApplication
        $state.ClosurePending = $false
        $state.ActiveGroup = ''
        $state.ActiveApplication = ''
        Save-ManualCheckpoint $state $Path
    }
    return $state
}

function Complete-ManualRow {
    param($Check, [object[]]$Steps)
    $recorded = @($Steps | Where-Object Result -ne 'Skip')
    if ($recorded.Count -eq 0) { return }
    $Check.Result = if (@($Steps | Where-Object Result -eq 'Fail').Count -gt 0) { 'Fail' }
        elseif (@($Steps | Where-Object { $_.Result -ne 'Pass' }).Count -gt 0) { 'Manual' }
        else { 'Pass' }
    foreach ($step in $Steps) {
        $Check.Notes = Add-ManualText $Check.Notes 'Manual evidence: ' "$($step.Name): $($step.Result). $($step.Notes) $($step.Assistance)"
        if ($step.IssueReference) { $Check.IssueReference = Add-ManualText $Check.IssueReference '' $step.IssueReference }
    }
    $Check.AutomationReview = Add-ManualText $Check.AutomationReview 'Manual review: ' "Completed by $env:USERDOMAIN\$env:USERNAME on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
}

function Invoke-ManualReview {
    param($State, [string]$CheckpointPath, [switch]$NoAssistedLaunch)
    foreach ($group in @(Get-ManualReviewGroups $State.Rows)) {
        $pending = @($group.Tasks | Where-Object { $_.Key -notin @($State.Steps | ForEach-Object { $_.Key }) })
        if ($pending.Count -eq 0) { continue }
        if ($State.ClosurePending) { throw 'Previous test-window closure must be confirmed before another application group can start.' }
        Write-Host "`n$($group.Sheet) / $($group.Name): $($pending.Count) remaining checks" -ForegroundColor Cyan
        Write-Host 'Keep this application open until this group is complete. Required close/reopen, restart or logoff steps inside a test remain operator-controlled.'
        $State.ActiveGroup = $group.Key
        $State.ActiveApplication = $group.Name
        $State.ClosurePending = $true
        Save-ManualCheckpoint $State $CheckpointPath
        $assistance = 'Operator-controlled application; no launch requested.'
        if ($NoAssistedLaunch) { Write-Host "Open $($group.Name) yourself for this group." }
        else {
            try {
                Start-ManualApplication $pending[0].Action
                $assistance = 'Application launch requested where available; startup and functional behavior are operator-confirmed. No processes are closed by the tool.'
            }
            catch {
                $assistance = "Assisted launch failed: $($_.Exception.Message)"
                Write-Warning $assistance
                $choice = Read-ManualChoice '[M]open the application yourself and continue, or [Q]save and exit' @('M', 'Q')
                if ($choice -eq 'Q') { throw 'Manual review stopped after a launch error. Progress was saved.' }
                $assistance += ' Operator chose manual launch; verify visible startup before recording Pass.'
            }
        }
        foreach ($task in $pending) {
            $check = $task.Check
            Write-Host "`n$($check.Sheet) / $($check.ID): $($check.Test) [$($task.Action.Name)]" -ForegroundColor Cyan
            Write-Host "Expected: $($check.Expected)"
            Write-Host "Automated evidence / saved notes: $($check.Notes)" -ForegroundColor DarkCyan
            if ($task.Action.Instructions) { Write-Host $task.Action.Instructions }
            $observation = Read-ManualObservation
            $State.Steps += @{
                Key = $task.Key; RowKey = $task.RowKey; Name = $task.Action.Name
                Result = $observation.Result; Notes = $observation.Notes
                IssueReference = $observation.IssueReference; Assistance = $assistance
            }
            $steps = @($State.Steps | Where-Object RowKey -eq $task.RowKey)
            if ($steps.Count -eq $task.StepCount -and $task.RowKey -notin @($State.CompletedRows)) {
                Complete-ManualRow $check $steps
                $State.CompletedRows += $task.RowKey
            }
            Save-ManualCheckpoint $State $CheckpointPath
            Write-Host 'Result and comment saved. Continuing to the next check.'
        }
        Confirm-ManualWindowClosure $group.Name
        $State.ClosurePending = $false
        $State.ActiveGroup = ''
        $State.ActiveApplication = ''
        Save-ManualCheckpoint $State $CheckpointPath
    }
}

