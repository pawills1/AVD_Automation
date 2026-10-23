[CmdletBinding()]
param
(
    [Parameter(Mandatory = $false)]
    [switch]
    $Install = $true,
    [Parameter(Mandatory = $false)]
    [string]$WorkDir = (Join-Path (Split-Path -Parent $PSScriptRoot) 'SelfHelp')
)

$ErrorActionPreference = 'Stop'

#//if $workdir specified, sets that as working location (if it exists). Otherwise, assumes script location is working directory.
If($workDir){

    If(Test-Path -Path "$($workdir)"){
        $origLocation = (Get-Location).path
        Set-Location -Path $($workDir)
    }else{
        throw "The working path '$workDir' does not exist."
    }

}

if ($install) {

    #determine installation log path;
    If (test-path -path "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs"){
        $logFolder =  "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs"
    }else{
        $logFolder =  "$env:windir\Logs"
    }

    $appNameForLog = $($workDir).split('\')[3]
    $LogPath = "$($logFolder)\$($appNameForLog)_AVD_Install.log"

    Start-Transcript -Path $logPath

    Write-Output "Creating C:\Windows\SySWOW64\ArmyJEA, C:\Program Files\WindowsPowerShell\Modules\JEA and C:\Program Files\SH folders..."
    mkdir "C:\Windows\SysWOW64\ArmyJEA" -ErrorAction SilentlyContinue
    mkdir "C:\Program Files\WindowsPowerShell\Modules\JEA" -ErrorAction SilentlyContinue
    mkdir "C:\Program Files\WindowsPowerShell\Modules\JEA\RoleCapabilities" -ErrorAction SilentlyContinue
    mkdir "C:\Program Files\SH" -ErrorAction SilentlyContinue
    mkdir "C:\ProgramData\UserHelp" -ErrorAction SilentlyContinue

    Write-Output "copying files to C:\Windows\SySWOW64\ArmyJEA, C:\Program Files\WindowsPowerShell\Modules\JEA and C:\program files\SH folders..."
    copy-item -Path "$($workdir)\User_Actions.psrc" -Destination "C:\Program Files\WindowsPowerShell\Modules\JEA\RoleCapabilities" -Force
    copy-item -Path "$($workDir)\User_Actions_Session.pssc" -Destination "C:\Program Files\WindowsPowerShell\Modules\JEA" -Force
    copy-item -Path "$($workDir)\Teams_ClearCache.ps1" -Destination "C:\Windows\SysWOW64\ArmyJEA" -Force
    copy-item -Path "$($workDir)\Teams_ClearCache-NewApp.ps1" -Destination "C:\Windows\SysWOW64\ArmyJEA" -Force
    copy-item -Path "$($workDir)\confirm_clearteams.ps1" -Destination "C:\Windows\SysWOW64\ArmyJEA" -Force
    copy-item -Path "$($workDir)\UserHelp.json" -Destination "C:\Program Files\SH" -Force
    copy-item -Path "$($workDir)\UserHelp.json" -Destination "C:\ProgramData\UserHelp" -force

    Write-Output "Registering Army JEA PS module..."
    $existingConfiguration = Get-PSSessionConfiguration -Name User_Actions -ErrorAction SilentlyContinue
    if ($null -ne $existingConfiguration) {
        Unregister-PSSessionConfiguration -Name User_Actions -Force
    }
    Register-PSSessionConfiguration -Name User_Actions -Path 'C:\Program Files\WindowsPowerShell\Modules\JEA\User_Actions_Session.pssc' -Force

    Write-Output "extracting contents of userhelp.zip to C:\program files\SH folder..."
    Expand-Archive "$($workDir)\userhelp.zip" -DestinationPath "C:\Program Files\SH" -force

    Write-Output "Creating Self-Help desktop shortcut..."
    $TargetFile = "C:\Program Files\SH\Appy.exe"
    $ShortcutFile = "C:\Users\Public\Desktop\SelfHelp.lnk"
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
    $Shortcut.TargetPath = $TargetFile
    $Shortcut.WorkingDirectory = 'C:\Program Files\SH'
    $Shortcut.Save()

    $startMenuShortcutFile = 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\SelfHelp.lnk'
    $startMenuShortcut = $WScriptShell.CreateShortcut($startMenuShortcutFile)
    $startMenuShortcut.TargetPath = $TargetFile
    $startMenuShortcut.WorkingDirectory = 'C:\Program Files\SH'
    $startMenuShortcut.Save()

    Write-Output "modifying ACL on C:\ProgramData\UserHelp folder..."
    $acl = Get-Acl -Path "C:\ProgramData\UserHelp"
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Users", "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
    $acl.SetAccessRule($accessRule)
    Set-Acl -Path "C:\ProgramData\UserHelp" -AclObject $acl

    if (-not (Test-Path -LiteralPath $TargetFile -PathType Leaf) -or -not (Test-Path -LiteralPath $ShortcutFile -PathType Leaf)) {
        throw 'Self-Help deployment verification failed.'
    }

    Write-Output "Installation of Self-Help tool completed successfully!"

    Stop-Transcript
}

#//change path back to location before $workdir set;
If($workDir){
    Set-Location -Path $origLocation
}
