[CmdletBinding()]
param([string]$SourceDirectory, [switch]$SkipCompiled)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
if (-not $SourceDirectory) { $SourceDirectory = Split-Path -Parent $PSScriptRoot }
$helper = Join-Path $SourceDirectory 'AVD-UAT-ManualActions.ps1'
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($helper, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw "Manual helper parse errors: $parseErrors" }
. $helper

$testRoot = Join-Path $env:TEMP ('AVD-Manual-AssistedTests-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot | Out-Null
$workbook = Join-Path $testRoot 'workbook.fixture'
$csv = Join-Path $testRoot 'evidence.csv'
Set-Content -LiteralPath $workbook -Value 'unchanged automated workbook fixture'
Set-Content -LiteralPath $csv -Value 'unchanged automated evidence fixture'
$script:passed = 0
$originalUserProfile = $env:USERPROFILE
$originalLocalAppData = $env:LOCALAPPDATA

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
    $script:passed++
}

function Assert-Throws {
    param([scriptblock]$Action, [string]$Pattern)
    $caught = $null
    try { & $Action }
    catch { $caught = $_ }
    Assert-True ($null -ne $caught -and $caught.Exception.Message -like $Pattern) "Expected '$Pattern'; received '$caught'."
}

function Reset-Fixture {
    $script:events = [System.Collections.Generic.List[string]]::new()
    $script:answers = [System.Collections.Generic.Queue[string]]::new()
    $script:available = $true
    $script:startFails = $false
    $fixtureProgramFiles = if ($env:ProgramW6432) { $env:ProgramW6432 } else { $env:ProgramFiles }
    $script:registeredPackages = @(
        @{ Name = 'Microsoft.Windows.Photos'; Version = '2608.1401.18.0'; InstallLocation = "$fixtureProgramFiles\WindowsApps\Microsoft.Windows.Photos_2608.1401.18.0_x64__fixture" }
    )
    $script:packageDirectories = @()
    $script:missingPaths = @()
    $script:packageQueryFails = $false
    $script:packageScanFails = $false
    $script:checkpoint = Join-Path $testRoot ([guid]::NewGuid().ToString('N') + '.json')
}

function Test-Path {
    param([string]$LiteralPath, [string]$PathType)
    if ($PathType -eq 'Leaf' -and $LiteralPath -notlike '*\tools\AvdManualFileTool\AvdManualFileTool.exe' -and
        ($LiteralPath -like '*.exe' -or $LiteralPath -like '*.lnk' -or $LiteralPath -like '*.msc')) {
        return $script:available -and $LiteralPath -notin $script:missingPaths
    }
    Microsoft.PowerShell.Management\Test-Path @PSBoundParameters
}

function Get-AppxPackage {
    [CmdletBinding()]
    param([string]$Name)
    if ($script:packageQueryFails) { throw 'Simulated package query failure' }
    $script:registeredPackages | Where-Object Name -eq $Name
}

function Get-ChildItem {
    [CmdletBinding()]
    param([string]$LiteralPath, [string]$Filter, [switch]$Directory, [switch]$File, [switch]$Recurse)
    if ($LiteralPath -like '*\WindowsApps') {
        if ($script:packageScanFails) { throw 'Simulated package directory access denied' }
        return $script:packageDirectories | Where-Object Name -like $Filter
    }
    Microsoft.PowerShell.Management\Get-ChildItem @PSBoundParameters
}

function Start-Process {
    param([string]$FilePath, [string[]]$ArgumentList, [string]$ErrorAction)
    $script:events.Add("open:$FilePath")
    if ($script:startFails) { throw 'Simulated launch failure' }
}

function Stop-Process { throw 'Manual workflow must NEVER stop a process.' }
function Get-Process { throw 'Manual workflow must not track or wait for PIDs.' }
function Get-CimInstance { throw 'Manual workflow must not enumerate process trees.' }

function Read-Host {
    param([string]$Prompt)
    if ($script:answers.Count -eq 0) { throw "Unexpected operator prompt: $Prompt" }
    if ($Prompt -like '*onfirm test windows*') {
        $saved = Get-Content -LiteralPath $script:checkpoint -Raw | ConvertFrom-Json
        Assert-True ($saved.ClosurePending) 'Closure must remain pending until confirmation is persisted.'
    }
    $answer = $script:answers.Dequeue()
    $script:events.Add("input:$answer")
    return $answer
}

function New-FixtureRow {
    param([string]$Id, [string]$Result = 'Manual', [string]$Sheet = 'Azure Virtual Desktop Tasks')
    return [pscustomobject]@{
        Sheet = $Sheet; ID = $Id; Category = 'Fixture'; Test = 'Fixture test'
        Expected = 'Fixture expected behavior'; Result = $Result; Notes = 'Original automated evidence'
        IssueReference = ''; AutomationReview = 'Original review'
    }
}

function Add-ReviewAnswers {
    param([object[]]$Rows, [string[]]$PassIds = @())
    foreach ($group in @(Get-ManualReviewGroups $Rows)) {
        foreach ($task in $group.Tasks) {
            if ($task.Check.ID -in $PassIds) {
                $script:answers.Enqueue('P')
                $script:answers.Enqueue("Verified $($task.Check.ID) in $($task.Action.Name)")
            } else { $script:answers.Enqueue('S') }
        }
        $script:answers.Enqueue('C')
    }
}

try {
    Reset-Fixture
    $programFiles = if ($env:ProgramW6432) { $env:ProgramW6432 } else { $env:ProgramFiles }
    $windowsApps = Join-Path $programFiles 'WindowsApps'
    $oldSnip = Join-Path $windowsApps 'Microsoft.ScreenSketch_11.9.0.0_x64__fixture'
    $newSnip = Join-Path $windowsApps 'Microsoft.ScreenSketch_11.10.0.0_x64__fixture'
    $paint = Join-Path $windowsApps 'Microsoft.Paint_11.10.0.0_x64__fixture'
    $photos = Join-Path $windowsApps 'Microsoft.Windows.Photos_2608.1401.18.0_x64__fixture'
    $script:registeredPackages = @(
        [pscustomobject]@{ Name = 'Microsoft.ScreenSketch'; Version = '11.9.0.0'; InstallLocation = $oldSnip }
        [pscustomobject]@{ Name = 'Microsoft.ScreenSketch'; Version = '11.10.0.0'; InstallLocation = $newSnip }
        [pscustomobject]@{ Name = 'Microsoft.Paint'; Version = '11.10.0.0'; InstallLocation = $paint }
        [pscustomobject]@{ Name = 'Microsoft.Windows.Photos'; Version = '2608.1401.18.0'; InstallLocation = $photos }
        [pscustomobject]@{ Name = 'Microsoft.Windows.Photos'; Version = '2608.999.0.0'; InstallLocation = "$windowsApps\Microsoft.Windows.Photos_2608.999.0.0_x64__fixture" }
        [pscustomobject]@{ Name = 'Microsoft.OneDrive.Photos'; Version = '9999.0.0.0'; InstallLocation = "$windowsApps\Microsoft.OneDrive.Photos_9999.0.0.0_x64__fixture" }
    )
    $photosAction = @(Get-ManualActions 'Azure Virtual Desktop Tasks' 'AVD-64' | Where-Object Name -eq 'Photos')[0]
    Start-ManualApplication $photosAction
    Assert-True ($script:events[$script:events.Count - 1] -eq "open:$photos\Photos.exe") 'Photos did not target the newest registered Windows Photos rather than OneDrive Photos.'
    foreach ($id in @('AVD-54', 'AVD-55', 'AVD-64')) {
        $action = @(Get-ManualActions 'Azure Virtual Desktop Tasks' $id | Where-Object Name -eq 'Snipping Tool')[0]
        Start-ManualApplication $action
        Assert-True ($script:events[$script:events.Count - 1] -eq "open:$newSnip\SnippingTool\SnippingTool.exe") "$id did not prefer the newest registered Snipping Tool package."
    }
    $paintAction = @(Get-ManualActions 'Azure Virtual Desktop Tasks' 'AVD-64' | Where-Object Name -eq 'Paint')[0]
    Start-ManualApplication $paintAction
    Assert-True ($script:events[$script:events.Count - 1] -eq "open:$paint\PaintApp\mspaint.exe") 'Paint did not prefer its packaged executable.'
    $script:missingPaths = @("$paint\PaintApp\mspaint.exe", "$newSnip\SnippingTool\SnippingTool.exe")
    Start-ManualApplication $action
    Assert-True ($script:events[$script:events.Count - 1] -eq "open:$newSnip\SnippingTool.exe") 'Root-level packaged Snipping Tool layout was not supported.'
    Start-ManualApplication $paintAction
    Assert-True ($script:events[$script:events.Count - 1] -eq "open:$paint\Paint\mspaint.exe") 'Older packaged Paint subfolder layout was not supported.'
    $script:missingPaths += "$paint\Paint\mspaint.exe"
    Start-ManualApplication $paintAction
    Assert-True ($script:events[$script:events.Count - 1] -eq "open:$paint\mspaint.exe") 'Root-level packaged Paint layout was not supported.'
    $script:registeredPackages = @()
    $script:packageDirectories = @($oldSnip, $newSnip, $paint | ForEach-Object {
        [pscustomobject]@{ Name = Split-Path -Leaf $_; FullName = $_ }
    })
    Start-ManualApplication $action
    Assert-True ($script:events[$script:events.Count - 1] -eq "open:$newSnip\SnippingTool.exe") 'Direct WindowsApps discovery did not sort versions numerically.'
    $script:missingPaths += "$newSnip\SnippingTool.exe"
    Start-ManualApplication $action
    Assert-True ($script:events[$script:events.Count - 1] -eq "open:$oldSnip\SnippingTool\SnippingTool.exe") 'Missing executable prevented discovery of another installed version.'
    $script:packageDirectories = @()
    Assert-Throws { Start-ManualApplication $photosAction } '*Microsoft.Windows.Photos*not found*'
    foreach ($fallback in @($action, $paintAction)) {
        Start-ManualApplication $fallback
        Assert-True ($script:events[$script:events.Count - 1] -eq "open:$($fallback.Path)") 'Missing packages did not fall back to System32.'
    }
    $script:packageQueryFails = $true
    $script:packageScanFails = $true
    $warnings = @()
    $fallbackPath = Resolve-ManualWindowsAppPath 'Paint' $paintAction.Path -WarningVariable warnings -WarningAction SilentlyContinue
    Assert-True ($fallbackPath -eq $paintAction.Path -and $warnings.Count -eq 2) 'Package discovery errors were not surfaced before fallback.'

    Reset-Fixture
    $hashFixture = Join-Path $testRoot 'hash.fixture'
    [IO.File]::WriteAllBytes($hashFixture, [byte[]]@(97, 98, 99))
    Assert-True ((Get-ManualFileHash $hashFixture) -ceq 'BA7816BF8F01CFEA414140DE5DAE2223B00361A396177A9CB410FF61F20015AD') 'Checkpoint SHA-256 does not match the known binary digest.'
    $expectedWorkbookHash = (Get-FileHash -LiteralPath $workbook -Algorithm SHA256).Hash
    $expectedCsvHash = (Get-FileHash -LiteralPath $csv -Algorithm SHA256).Hash
    function Get-FileHash { throw 'Get-FileHash is unavailable in this host.' }
    try {
        $hashRows = @((New-FixtureRow 'AVD-03'))
        $hashState = New-ManualCheckpoint $workbook $csv $hashRows
        Assert-True ($hashState.WorkbookHash -ceq $expectedWorkbookHash -and $hashState.CsvHash -ceq $expectedCsvHash) 'Hash format is incompatible with existing checkpoints.'
        Save-ManualCheckpoint $hashState $script:checkpoint
        $resumed = Read-ManualCheckpoint $script:checkpoint $workbook $csv $hashRows
        Assert-True ($resumed.WorkbookHash -ceq $expectedWorkbookHash) 'Checkpoint resume requires Get-FileHash.'
        $previousCheckpoint = Get-Content -LiteralPath $script:checkpoint -Raw
        $lockedCheckpoint = [IO.File]::Open($script:checkpoint, 'Open', 'ReadWrite', 'None')
        try {
            $hashState.ActiveApplication = 'Unsaved fixture'
            Assert-Throws { Save-ManualCheckpoint $hashState $script:checkpoint } '*Manual file helper failed*'
        }
        finally { $lockedCheckpoint.Dispose() }
        Assert-True ((Get-Content -LiteralPath $script:checkpoint -Raw) -ceq $previousCheckpoint) 'Failed checkpoint replacement damaged saved progress.'
        Assert-True (@(Get-ChildItem -LiteralPath $testRoot -Filter '*.tmp').Count -eq 0) 'Failed checkpoint replacement leaked a temporary file.'
        Assert-Throws { Get-ManualFileHash (Join-Path $testRoot 'missing.fixture') } '*Manual file helper failed*'
        $exclusive = [IO.File]::Open($workbook, 'Open', 'ReadWrite', 'None')
        $exclusive.Dispose()
    }
    finally { Remove-Item Function:\Get-FileHash }
    Reset-Fixture
    $commands = @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.CommandAst] }, $true) |
        ForEach-Object { $_.GetCommandName() })
    Assert-True (@($commands | Where-Object { $_ -in @('Stop-Process', 'taskkill', 'Get-Process', 'Get-CimInstance') }).Count -eq 0) 'Manual helper contains PID management.'
    Assert-True ($ast.Extent.Text -notmatch '\.(Kill|CloseMainWindow|WaitForExit)\s*\(') 'Manual helper contains automatic close/exit-wait logic.'
    foreach ($mapping in @(
        @('AVD-05', 'Adobe Acrobat'), @('AVD-06', 'AWS WickrGov'), @('AVD-08', 'DAX Studio'),
        @('AVD-10', 'DTP Lite'), @('AVD-18', 'Microsoft Copilot Chat'),
        @('AVD-41', 'Outlook'), @('AVD-42', 'Outlook'), @('AVD-48', 'Edge'),
        @('AVD-53', 'SelfHelp'), @('AVD-54', 'Snipping Tool'), @('AVD-55', 'Snipping Tool'),
        @('AVD-56', 'STIG Viewer'), @('AVD-57', 'YubiKey Manager'),
        @('AVD-58', 'Edge'), @('AVD-61', 'File Explorer'), @('AVD-62', 'File Explorer'),
        @('AVD-69', 'OneDrive'), @('AVD-70', 'OneDrive'), @('AVD-74', 'Windows Update Settings')
    )) {
        Assert-True (@(Get-ManualActions 'Azure Virtual Desktop Tasks' $mapping[0])[0].Name -eq $mapping[1]) "Stale mapping for $($mapping[0])."
    }
    Assert-True (@(Get-ManualActions 'Azure Virtual Desktop Tasks' 'AVD-19').Count -eq 7) 'Office startup row must cover seven applications.'
    Assert-True (@(Get-ManualActions 'Azure Virtual Desktop Tasks' 'AVD-20').Count -eq 7) 'Office update row must cover seven applications.'
    Assert-True (@(Get-ManualActions 'Remote Apps Tasks' 'AVD-19').Count -eq 0) 'Worksheet mappings leaked.'
    foreach ($number in 2..17) {
        $remote = @(Get-ManualActions 'Remote Apps Tasks' ('AVD-{0:D2}' -f $number))
        Assert-True ($remote.Count -eq 1 -and -not $remote[0].Path -and $remote[0].Instructions -like '*Remote App*') 'Remote Apps review must not silently launch local substitutes.'
    }
    Assert-True (-not (Test-ManualReviewRequired (New-FixtureRow 'AVD-81' 'Pass'))) 'A completed machine-only check was selected.'

    $rows = @(
        New-FixtureRow 'AVD-03'
        New-FixtureRow 'AVD-19'
        New-FixtureRow 'AVD-04'
        New-FixtureRow 'AVD-41'
        New-FixtureRow 'AVD-20'
        New-FixtureRow 'AVD-42'
        New-FixtureRow 'AVD-16' 'Manual' 'Remote Apps Tasks'
    )
    $groups = @(Get-ManualReviewGroups $rows)
    Assert-True (@($groups | Where-Object Name -eq 'Adobe Acrobat').Count -eq 1) 'Acrobat checks were not grouped.'
    $outlook = @($groups | Where-Object { $_.Name -eq 'Outlook' -and $_.Sheet -eq 'Azure Virtual Desktop Tasks' })[0]
    Assert-True (($outlook.Tasks.Check.ID -join ',') -eq 'AVD-19,AVD-41,AVD-20,AVD-42') 'Outlook checks are not in the same application session.'
    Assert-True (@($groups | Where-Object Name -eq 'Outlook').Count -eq 2) 'Different worksheet environments were merged.'

    $state = New-ManualCheckpoint $workbook $csv $rows
    Add-ReviewAnswers $rows @('AVD-03', 'AVD-04', 'AVD-19', 'AVD-20', 'AVD-41', 'AVD-42', 'AVD-16')
    Invoke-ManualReview $state $script:checkpoint
    $opens = @($script:events | Where-Object { $_ -like 'open:*' })
    Assert-True ($opens.Count -eq $groups.Count - 1) 'An application was reopened within its group or a local Remote App substitute was launched.'
    Assert-True (@($script:events | Where-Object { $_ -like '*Acrobat.exe' }).Count -eq 1) 'Acrobat was reopened.'
    Assert-True (@($rows | Where-Object Result -ne 'Pass').Count -eq 0) 'Grouped results were not aggregated correctly.'
    Assert-True ($state.CompletedRows.Count -eq $rows.Count -and -not $state.ClosurePending) 'Rows or closure were not completed.'
    $saved = Get-Content -LiteralPath $script:checkpoint -Raw | ConvertFrom-Json
    Assert-True ($saved.Steps.Count -eq 19 -and $script:answers.Count -eq 0) 'Per-application observations were lost or extra prompts appeared.'
    foreach ($row in $rows) {
        Assert-True ($row.Notes -like '*Original automated evidence*Verified*' -and $row.AutomationReview -like '*Original review*Manual review*') 'Automated evidence was discarded.'
    }

    Reset-Fixture
    $rows = @((New-FixtureRow 'AVD-03'), (New-FixtureRow 'AVD-04'), (New-FixtureRow 'AVD-54'))
    $state = New-ManualCheckpoint $workbook $csv $rows
    foreach ($answer in @('P', 'Acrobat first check saved', 'Q')) { $script:answers.Enqueue($answer) }
    Assert-Throws { Invoke-ManualReview $state $script:checkpoint } '*stopped by the operator*'
    $saved = Get-Content -LiteralPath $script:checkpoint -Raw | ConvertFrom-Json
    Assert-True ($saved.ClosurePending -and $saved.Steps.Count -eq 1 -and $saved.Rows[0].Result -eq 'Pass') 'Interruption lost the confirmed result or pending window state.'
    Assert-True (@($script:events | Where-Object { $_ -like 'open:*' }).Count -eq 1) 'Another app launched after cancellation.'
    $script:answers.Enqueue('Q')
    Assert-Throws { Read-ManualCheckpoint $script:checkpoint $workbook $csv $rows } '*closure is unconfirmed*'
    $script:answers.Enqueue('C')
    $restored = Read-ManualCheckpoint $script:checkpoint $workbook $csv $rows
    foreach ($answer in @('P', 'Acrobat second check saved', 'C', 'P', 'Snip visible', 'C')) { $script:answers.Enqueue($answer) }
    Invoke-ManualReview $restored $script:checkpoint
    Assert-True ($restored.Steps.Count -eq 3 -and $restored.Rows[0].Notes -notlike '*second check*') 'Resume duplicated or misattributed evidence.'
    $eventCount = $script:events.Count
    Invoke-ManualReview $restored $script:checkpoint
    Assert-True ($script:events.Count -eq $eventCount) 'Completed review reopened applications.'

    Reset-Fixture
    $row = New-FixtureRow 'AVD-03'
    $state = New-ManualCheckpoint $workbook $csv @($row)
    foreach ($answer in @('F', 'Failure evidence saved', 'ISSUE-123', 'Q')) { $script:answers.Enqueue($answer) }
    Assert-Throws { Invoke-ManualReview $state $script:checkpoint } '*closure is unconfirmed*'
    $saved = Get-Content -LiteralPath $script:checkpoint -Raw | ConvertFrom-Json
    Assert-True ($saved.Rows[0].Result -eq 'Fail' -and $saved.Rows[0].IssueReference -eq 'ISSUE-123' -and $saved.ClosurePending) 'Failure evidence was not saved before closure confirmation.'
    $script:answers.Enqueue('C')
    $restored = Read-ManualCheckpoint $script:checkpoint $workbook $csv @($row)
    $eventCount = $script:events.Count
    Invoke-ManualReview $restored $script:checkpoint
    Assert-True ($script:events.Count -eq $eventCount) 'A group completed before interruption was reopened on resume.'
    $restored.WorkbookHash = 'incorrect'
    Save-ManualCheckpoint $restored $script:checkpoint
    Assert-Throws { Read-ManualCheckpoint $script:checkpoint $workbook $csv @($row) } '*WorkbookHash does not match*'

    Reset-Fixture
    $row = New-FixtureRow 'AVD-03'
    $state = New-ManualCheckpoint $workbook $csv @($row)
    $script:available = $false
    foreach ($answer in @('M', 'P', 'Manually launched after missing executable', 'C')) { $script:answers.Enqueue($answer) }
    Invoke-ManualReview $state $script:checkpoint
    Assert-True ($row.Result -eq 'Pass' -and $row.Notes -like '*executable not found*Operator chose manual launch*') 'Missing application was silently passed or manual recovery was lost.'
    Reset-Fixture
    $row = New-FixtureRow 'AVD-03'
    $state = New-ManualCheckpoint $workbook $csv @($row)
    $script:startFails = $true
    $script:answers.Enqueue('Q')
    Assert-Throws { Invoke-ManualReview $state $script:checkpoint } '*stopped after a launch error*'
    Assert-True ($state.Steps.Count -eq 0 -and $state.ClosurePending) 'Launch error fabricated evidence or cleared closure state.'

    Reset-Fixture
    $row = New-FixtureRow 'AVD-54'
    $state = New-ManualCheckpoint $workbook $csv @($row)
    foreach ($answer in @('P', '', 'Visible snip', 'C')) { $script:answers.Enqueue($answer) }
    Invoke-ManualReview $state $script:checkpoint -NoAssistedLaunch
    Assert-True (@($script:events | Where-Object { $_ -like 'open:*' }).Count -eq 0 -and $row.Notes -like '*Visible snip*') 'Guided launch mode launched an application or accepted an empty comment.'
    Reset-Fixture
    $state = New-ManualCheckpoint $workbook $csv @((New-FixtureRow 'AVD-03'))
    Assert-Throws { Invoke-ManualReview $state (Join-Path $testRoot 'missing\checkpoint.json') } '*Could not find a part of the path*'
    Assert-True ($script:events.Count -eq 0) 'Application launched despite a checkpoint write failure.'

    foreach ($results in @(@('Skip', 'Skip'), @('Pass', 'Skip'), @('Pass', 'Fail'), @('Pass', 'Pass'))) {
        $row = New-FixtureRow 'AVD-19' 'Partial Pass'
        $steps = @($results | ForEach-Object { [pscustomobject]@{ Result = $_; Name = 'App'; Notes = 'Evidence'; IssueReference = ''; Assistance = '' } })
        Complete-ManualRow $row $steps
        $expected = if ($results -contains 'Fail') { 'Fail' } elseif ($results -notcontains 'Pass') { 'Partial Pass' } elseif ($results -contains 'Skip') { 'Manual' } else { 'Pass' }
        Assert-True ($row.Result -eq $expected) "Composite result $($results -join ',') was incorrect."
    }

    Reset-Fixture
    $deployment = Join-Path $testRoot 'deployment'
    $automatedDirectory = Join-Path $deployment 'Automated_Results_20261111-110000-000'
    New-Item -ItemType Directory -Path $automatedDirectory -Force | Out-Null
    foreach ($name in @('AVD-UAT-Checklist-Manual.ps1', 'AVD-UAT-ManualActions.ps1', 'AVD-UAT-Paths.ps1')) {
        Copy-Item -LiteralPath (Join-Path $SourceDirectory $name) -Destination $deployment
    }
    $fileToolDirectory = Join-Path $deployment 'tools\AvdManualFileTool'
    New-Item -ItemType Directory -Path $fileToolDirectory -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $SourceDirectory 'tools\AvdManualFileTool\AvdManualFileTool.exe') -Destination $fileToolDirectory
    Set-Content -LiteralPath (Join-Path $deployment 'AVD-UAT-Prerequisite-Check.ps1') -Value '@{ Passed = $true; LogPath = "fixture.log"; ReportPath = "fixture.html" }'
    $env:USERPROFILE = Join-Path $testRoot 'profile'
    $env:LOCALAPPDATA = Join-Path $testRoot 'local'
    $workbookTool = Join-Path $SourceDirectory 'tools\AvdChecklistTool\AvdChecklistTool.exe'
    $automatedWorkbook = Join-Path $automatedDirectory 'AVD Checklist - Completed - 20261111-110000-000.xlsx'
    $automatedCsv = Join-Path $automatedDirectory 'AVD Checklist Results - 20261111-110000-000.csv'
    $exportedCsv = Join-Path $testRoot 'exported-template.csv'
    Copy-Item -LiteralPath (Join-Path $SourceDirectory 'AVD Test Plan Template - DO NOT DELETE OR MOVE.xlsx') -Destination $automatedWorkbook
    & $workbookTool export --input $automatedWorkbook --output $exportedCsv | Out-Null
    Assert-True ($LASTEXITCODE -eq 0) 'Template export failed.'
    $evidence = @(Import-Csv -LiteralPath $exportedCsv | Select-Object Sheet, ID, Category, Test,
        @{ Name = 'Expected'; Expression = { $_.'Expected Results' } },
        @{ Name = 'ValidationCheck'; Expression = { 'Fixture validation' } },
        @{ Name = 'Result'; Expression = {
            if ($_.ID -eq 'AVD-03') { 'Manual' } elseif ($_.ID -eq 'AVD-38') { 'Fail' } else { 'Pass' }
        } },
        @{ Name = 'IssueReference'; Expression = { '' } },
        @{ Name = 'Notes'; Expression = { 'Fixture automated evidence' } },
        @{ Name = 'AutomationReview'; Expression = { 'Fixture automated review' } })
    $evidence | Export-Csv -LiteralPath $automatedCsv -NoTypeInformation -Encoding UTF8
    $sourceHash = (Get-FileHash -LiteralPath $automatedWorkbook).Hash
    $csvHash = (Get-FileHash -LiteralPath $automatedCsv).Hash
    $fixtureLauncher = Join-Path $deployment 'AVD-UAT-Checklist-Manual.ps1'
    $null = . $fixtureLauncher -UatDirectory $deployment -WorkbookTool $workbookTool -SkipUnresolved
    Assert-True ($script:events.Count -eq 0) '-SkipUnresolved launched applications or prompted.'
    $documents = Join-Path $env:USERPROFILE 'OneDrive - US Army\Documents'
    $skipCsv = @(Get-ChildItem -LiteralPath $documents -Filter 'AVD Checklist Final Results*.csv')[0]
    Assert-True (@(Import-Csv -LiteralPath $skipCsv.FullName | Where-Object Result -eq 'Manual').Count -eq 2) 'Skip workflow changed evidence.'
    Add-ReviewAnswers $evidence @('AVD-03', 'AVD-04', 'AVD-05', 'AVD-19', 'AVD-20', 'AVD-41', 'AVD-42', 'AVD-38', 'AVD-54', 'AVD-55')
    $null = . $fixtureLauncher -UatDirectory $deployment -WorkbookTool $workbookTool -CheckpointPath $script:checkpoint
    Assert-True ($script:answers.Count -eq 0) 'Launcher did not execute the expected grouped prompts.'
    $finalCsv = @(Get-ChildItem -LiteralPath $documents -Filter 'AVD Checklist Final Results*.csv' | Sort-Object Name -Descending)[0]
    $finalRows = @(Import-Csv -LiteralPath $finalCsv.FullName)
    Assert-True ($finalRows.Count -eq 99) 'Final CSV changed row count.'
    foreach ($id in @('AVD-03', 'AVD-04', 'AVD-05', 'AVD-19', 'AVD-20', 'AVD-41', 'AVD-42', 'AVD-38', 'AVD-54', 'AVD-55')) {
        $reviewed = @($finalRows | Where-Object { $_.Sheet -eq 'Azure Virtual Desktop Tasks' -and $_.ID -eq $id })[0]
        Assert-True ($reviewed.Result -eq 'Pass' -and $reviewed.Notes -like "*Fixture automated evidence*Verified $id*") "$id result/comment missing from final CSV."
    }
    $books = @(Get-ChildItem -LiteralPath $documents -Filter 'AVD Checklist - Final*.xlsx')
    $reports = @(Get-ChildItem -LiteralPath $documents -Filter 'AVD UAT Final Report*.html')
    Assert-True ($books.Count -eq 2 -and $reports.Count -eq 2) 'Final outputs were not generated.'
    $latestBook = @($books | Sort-Object Name -Descending)[0]
    & $workbookTool export --input $latestBook.FullName --output $exportedCsv | Out-Null
    Assert-True ($LASTEXITCODE -eq 0) 'Final workbook could not be reopened.'
    $bookRow = @(Import-Csv -LiteralPath $exportedCsv | Where-Object { $_.Sheet -eq 'Azure Virtual Desktop Tasks' -and $_.ID -eq 'AVD-19' })[0]
    Assert-True ($bookRow.Result -eq 'Pass' -and $bookRow.Notes -like '*Excel*Verified AVD-19*') 'Workbook lost composite Office results/comments.'
    $html = Get-Content -LiteralPath (@($reports | Sort-Object Name -Descending)[0].FullName) -Raw
    Assert-True ($html -like '*Verified AVD-54*' -and $html -like '*Verified AVD-41*') 'HTML lost manual evidence.'
    $saved = Get-Content -LiteralPath $script:checkpoint -Raw | ConvertFrom-Json
    Assert-True ($saved.Outputs.Count -eq 3) 'Checkpoint was discarded or finalized before all three outputs succeeded.'
    Assert-True (@($script:events | Where-Object { $_ -like '*Acrobat.exe' }).Count -eq 1) 'Acrobat should launch once, not once per check.'
    $eventCount = $script:events.Count
    $null = . $fixtureLauncher -UatDirectory $deployment -WorkbookTool $workbookTool -CheckpointPath $script:checkpoint
    Assert-True ($script:events.Count -eq $eventCount) 'Report regeneration reopened apps or asked for results again.'
    Assert-True ((Get-FileHash -LiteralPath $automatedWorkbook).Hash -eq $sourceHash -and (Get-FileHash -LiteralPath $automatedCsv).Hash -eq $csvHash) 'Manual flow modified automated source evidence.'

    $failureCheckpoint = Join-Path $testRoot 'report-retry.json'
    $failureTool = Join-Path $testRoot 'failed-workbook-tool.cmd'
    Set-Content -LiteralPath $failureTool -Value '@exit /b 9' -Encoding ASCII
    Add-ReviewAnswers $evidence
    $script:checkpoint = $failureCheckpoint
    Assert-Throws {
        $null = . $fixtureLauncher -UatDirectory $deployment -WorkbookTool $failureTool -CheckpointPath $failureCheckpoint
    } '*Workbook update failed with exit code 9*'
    $failedState = Get-Content -LiteralPath $failureCheckpoint -Raw | ConvertFrom-Json
    Assert-True ($failedState.Steps.Count -gt 0 -and $failedState.Outputs.Count -eq 0) 'Export failure discarded progress or falsely recorded completed exports.'
    $eventCount = $script:events.Count
    $null = . $fixtureLauncher -UatDirectory $deployment -WorkbookTool $workbookTool -CheckpointPath $failureCheckpoint
    Assert-True ($script:events.Count -eq $eventCount) 'Report retry repeated prompts or reopened applications.'
    $recoveredState = Get-Content -LiteralPath $failureCheckpoint -Raw | ConvertFrom-Json
    Assert-True ($recoveredState.Outputs.Count -eq 3) 'Report retry did not persist its complete output set.'
    Assert-True (@(Get-ChildItem -LiteralPath $testRoot -Recurse -Filter '*.tmp').Count -eq 0) 'Checkpoint writes leaked temporary files.'

    if (-not $SkipCompiled) {
        $fixtureExe = Join-Path $deployment 'AVD-UAT-Checklist-Manual.exe'
        Copy-Item -LiteralPath (Join-Path $SourceDirectory 'AVD-UAT-Checklist-Manual.exe') -Destination $fixtureExe
        & $fixtureExe -UatDirectory $deployment -WorkbookTool $workbookTool -SkipUnresolved | Out-Null
        Assert-True ($LASTEXITCODE -eq 0) 'Compiled skip workflow failed.'
        $compiledFixtures = @'
function Get-FileHash { throw 'Get-FileHash is unavailable in this compiled host.' }
function Start-ManualApplication {
    param($Action)
    Add-Content -LiteralPath (Join-Path $PSScriptRoot 'compiled-events.log') -Value "open:$($Action.Name)"
}
function Read-ManualObservation {
    Add-Content -LiteralPath (Join-Path $PSScriptRoot 'compiled-events.log') -Value 'observation'
    return @{ Result = 'Pass'; Notes = 'Compiled fixture evidence'; IssueReference = '' }
}
function Confirm-ManualWindowClosure {
    param($Name)
    Add-Content -LiteralPath (Join-Path $PSScriptRoot 'compiled-events.log') -Value "confirm:$Name"
}
'@
        $fixtureHelper = Join-Path $deployment 'AVD-UAT-ManualActions.ps1'
        $fixtureSource = (Get-Content -LiteralPath $fixtureHelper -Raw) -replace '(?ms)^# SIG # Begin signature block.*\z', ''
        Set-Content -LiteralPath $fixtureHelper -Value ($fixtureSource + "`r`n" + $compiledFixtures) -Encoding UTF8
        $compiledCheckpoint = Join-Path $testRoot 'compiled.json'
        & $fixtureExe -UatDirectory $deployment -WorkbookTool $workbookTool -CheckpointPath $compiledCheckpoint | Out-Null
        Assert-True ($LASTEXITCODE -eq 0) 'Compiled grouped workflow failed.'
        $compiledEvents = @(Get-Content -LiteralPath (Join-Path $deployment 'compiled-events.log'))
        $expectedEvents = @(
            foreach ($group in @(Get-ManualReviewGroups $evidence)) {
                "open:$($group.Name)"
                foreach ($task in $group.Tasks) { 'observation' }
                "confirm:$($group.Name)"
            }
        )
        Assert-True (($compiledEvents -join '|') -eq ($expectedEvents -join '|')) 'Compiled launcher did not group launch, observations and operator confirmation correctly.'
        $compiledState = Get-Content -LiteralPath $compiledCheckpoint -Raw | ConvertFrom-Json
        Assert-True ($compiledState.Outputs.Count -eq 3 -and -not $compiledState.ClosurePending) 'Compiled output checkpoint was incomplete.'
        foreach ($output in $compiledState.Outputs) { Assert-True (Test-Path -LiteralPath $output) 'Compiled output is missing.' }
        Assert-True ($compiledState.WorkbookHash -ceq $sourceHash -and $compiledState.CsvHash -ceq $csvHash) 'Compiled checkpoint hashes differ from Get-FileHash.'
        & $fixtureExe -UatDirectory $deployment -WorkbookTool $workbookTool -CheckpointPath $compiledCheckpoint | Out-Null
        Assert-True ($LASTEXITCODE -eq 0) 'Compiled checkpoint resume failed without Get-FileHash.'
        Assert-True (@(Get-Content -LiteralPath (Join-Path $deployment 'compiled-events.log')).Count -eq $compiledEvents.Count) 'Compiled resume repeated completed application checks.'
        $compiledRows = @(Import-Csv -LiteralPath $compiledState.Outputs[0])
        Assert-True (@($compiledRows | Where-Object Notes -like '*Compiled fixture evidence*').Count -gt 60) 'Compiled results did not reach final CSV.'
    }
    Write-Output "PASS: $script:passed assertions; current mappings, grouping, zero PID management, incremental saving, cancellation/resume, CSV/XLSX/HTML and automated-input preservation. App launches are mocked."
}
finally {
    $env:USERPROFILE = $originalUserProfile
    $env:LOCALAPPDATA = $originalLocalAppData
    Remove-Item -LiteralPath $testRoot -Recurse -Force
}
