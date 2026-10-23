<#
[CmdletBinding()]
param
(
    [Parameter(Mandatory = $false)]
    [switch]
    $Install,
    [Parameter(Mandatory = $false)]
    [switch]
    $Uninstall,
    [Parameter(Mandatory = $false)]
    [string]$workDir
)
#>

$ErrorActionPreference = 'Stop'
$install = $true
$softwareRoot = Split-Path -Parent $PSScriptRoot
$workDir = Join-Path $softwareRoot 'Notepad++'

#//if $workdir specified, sets that as working location (if it exists). Otherwise, assumes script location is working directory.
If($workDir){

    If(Test-Path -Path "$($workdir)"){
        $origLocation = (Get-Location).path
        Set-Location -Path $($workDir)
    }else{
        Write-Warning "The working path $($workDir) does not exit. Script will exit."
        exit
    }

}

if ($install) {

    #determine installation log path;
    If (test-path -path "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs"){
        $logPath = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\Notepad++_AVD_Install.log"
    }else{
        $logPath = "$env:windir\Logs\Notepad++_AVD_Install.log"
    }

    Start-Transcript -Path $logPath

    $installFiles = @(Get-ChildItem -LiteralPath $softwareRoot -Filter 'npp*.exe' -File)
    if (Test-Path -LiteralPath $workDir -PathType Container) {
        $installFiles += @(Get-ChildItem -LiteralPath $workDir -Filter 'npp*.exe' -File -Recurse)
    }
    $installFiles = @($installFiles | Sort-Object -Property FullName -Unique)
    if ($installFiles.Count -ne 1) {
        throw "Expected exactly one Notepad++ installer named 'npp*.exe' under '$softwareRoot'; found $($installFiles.Count). Add the offline installer payload and rerun."
    }

    Write-Output "Installing $($installFiles[0].FullName)..."
    $process = Start-Process -FilePath $installFiles[0].FullName -ArgumentList '/S' -Wait -PassThru
    if ($process.ExitCode -ne 0) {
        throw "Notepad++ installer returned exit code $($process.ExitCode)."
    }

    $installedExecutable = Join-Path $env:ProgramFiles 'Notepad++\notepad++.exe'
    if (-not (Test-Path -LiteralPath $installedExecutable -PathType Leaf)) {
        throw "Notepad++ installation verification failed: $installedExecutable"
    }

    Write-Output "Notepad++ installed successfully to '$installedExecutable'."

    Stop-Transcript
}
if ($Uninstall) {

    #determine installation log path;
    If (test-path -path "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs"){
        $logPath = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\Firefox_AVD_Uninstall.log"
    }else{
        $logPath = "$env:windir\Logs\Firefox_AVD_Unistall.log"
    }

    Start-Transcript -Path $logPath

    #Detect installer
    $installer = Get-childitem -path .
    If ($InstallerPath.count -gt 1) {
        #If more than one file found, filter for installer.
        Write-Output "More than one file, filtering for the installer."
        $installFile = $InstallerPath | Where-Object { $_.Extension -match ".exe" -or $_.Extension -match ".msi" }
    }
    else {
        #The install file is the only file in there.
        $installFile = $InstallerPath
    }

    if ($installFile.Extension -match ".msi") {
        #Use Get-MSIProperties to grab product code of MSI and begin uninstall
        function Get-MSIProperties {
            param (
                [Parameter(Mandatory = $true)]
                [ValidateNotNullOrEmpty()]
                [System.IO.FileInfo] $path,

                [string[]] $properties = @('ProductCode', 'ProductVersion', 'ProductName', 'Manufacturer', 'ProductLanguage')
            )
            begin {
                $windowsInstaller = (New-Object -ComObject WindowsInstaller.Installer)
            }
            process {
                $table = @{}
                $msi = $windowsInstaller.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $windowsInstaller, @($Path.FullName, 0))
                foreach ($property in $properties) {
                    try {
                        $view = $msi.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $msi, ("SELECT Value FROM Property WHERE Property = '$($property)'"))
                        $view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null)
                        $record = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)
                        $table.add($property, $record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, 1))
                    }
                    catch {
                        $table.add($property, $null)
                    }
                }
                $msi.GetType().InvokeMember('Commit', 'InvokeMethod', $null, $msi, $null)
                $view.GetType().InvokeMember('Close', 'InvokeMethod', $null, $view, $null)
                $msi = $null
                $view = $null
                return $table
            }
            end {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($windowsInstaller) | Out-Null
                Return [System.GC]::Collect()
            }
        }

        $properties = Get-MSIProperties -path $installFile.FullName
        $prodInfo = $properties.ProductCode

        #Begin Uninstall

        Write-Output "Uninstall String is '/X$prodInfo' /qn.  Beginning Uninstall.";
        Start-Process "msiexec.exe" -ArgumentList "/X$prodInfo /qn" -Wait

    }
    elseif ($installer.Extension -match ".exe") {
        #Firefox specific
        $uninstallString = "$env:ProgramW6432\Mozilla Firefox\uninstall\helper.exe"

        #Beginning Uninstallation
        start-process $uninstallString -ArgumentList "/s" -Wait

    }
    else {
        Write-Output "Unable to determine type."
    }

    Stop-Transcript
}

#//change path back to location before $workdir set;
If($workDir){
    Set-Location -Path $origLocation
}

