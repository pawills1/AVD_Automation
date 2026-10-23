$workDir = "C:\temp\bootstrap\GCCSelfHelp"
$useZIPfile = $true

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

#//determine installation log path;
If (!(test-path -path "$env:SystemDrive\Program Files\AVDAppInstallLogs")){
    New-Item -ItemType Directory -Path "$env:SystemDrive\Program Files\AVDAppInstallLogs" -Force
}

$logFolder = "$env:SystemDrive\Program Files\AVDAppInstallLogs"
$appNameForLog = $($workDir).split('\')[3]
$LogPath = "$($logFolder)\$($appNameForLog)_AVD_Install.log"

Start-Transcript -Path $logPath

$InstallerPath = Get-childitem -path . -File

    If($useZIPfile -eq $TRUE){
    $ZIPFile = $InstallerPath | Where-Object {$_.Extension -match ".zip"}
    Write-Output "Extracting contents of $($ZIPFile) to $($workDir)..."
    Expand-Archive -literalPath $ZIPFile -DestinationPath $workDir -Force
}

write-output "Begin setup and configuration of JEA Module..."

#//create $selfHelpFolderPath if not there already
$selfHelpFolderPath = "C:\Program Files\SelfHelp"

if (!(test-path -Path $selfHelpFolderPath)){
    Write-Output "$($selfHelpFolderPath) doesn't exist, creating it..."
    New-Item -ItemType Directory -Path $selfHelpFolderPath -Force | Out-Null
}

#//Copy each JSON and NLog file to the SelfHelp folder and verify if copy was successful
$jsonNlogFiles = @("UserHelp.json", "FAQ.json", "config.json", "NLog.config")

foreach ($file in $jsonNlogFiles) {
    Copy-Item -Path "$($workdir)\$($file)" -Destination $selfHelpFolderPath -Force | Out-Null

    If (Test-Path -Path "$($selfHelpFolderPath)\$($file)"){
        Write-Output "Successfully copied $($file) into $($selfHelpFolderPath)..."
    }Else{
        Write-Output "Could NOT copy $($file) into $($selfHelpFolderPath)..."
    }
    
}

#//Create registry key which SelfHelp.exe calls for to locate files
$regConfigPath = "HKLM:\SOFTWARE\AUDSSH"
Write-Output "Creating $($regConfigPath) key and value..."
New-Item -Path $regConfigPath -Force | Out-Null
Set-ItemProperty -Path $regConfigPath -Name ConfigPath -Value (Join-Path -Path $selfHelpFolderPath -ChildPath 'config.json') -Force | Out-Null
Write-Output "Created $($regConfigPath) registry key and set values."

#//create GCCJeaModule folder;
$jeaModulePath = "C:\Program Files\WindowsPowerShell\Modules\GCCJeaModule"

If (!(test-path -path $jeaModulePath)){
    write-output "$($jeaModulePath) doesn't exist, creating it..."
    New-Item -ItemType Directory -Path $jeaModulePath -Force | Out-Null
}

#//Create RoleCapabilities folder
$roleCapabilitiesPath = "$jeaModulePath\RoleCapabilities"

If (!(test-path -path $roleCapabilitiesPath)){
    write-output "$($roleCapabilitiesPath) doesn't exist, creating it..."
    New-Item -ItemType Directory -Path $roleCapabilitiesPath -Force | Out-Null
}

<#
#//Create ArmyJEA folder
$ArmyJEAPath = "C:\Windows\SysWOW64\ArmyJEA"

If (!(test-path -path $ArmyJEAPath)){
    write-output "$($ArmyJEAPath) doesn't exist, creating it..."
    New-Item -ItemType Directory -Path $ArmyJEAPath -Force | Out-Null
}
#>

#//copy OPSECMode.png to wallpaper location for system;
Copy-Item -Path "$($workdir)\OPSECMode.png" -Destination "C:\Windows\Web\Wallpaper\Windows\OPSECMode.png" -Force

#//copy JEA files to JEA directories & verify successful copy;
$jeaModuleFiles = @("GCCJeaEndpoint.pssc","CustomFunctions.ps1","InitScript.ps1","ToggleOPSECMode.ps1","confirm_clearteams.ps1","confirm_customstartmenu.ps1","Teams_ClearCache-NewApp.ps1","DeleteUserCache.ps1","confirm_DeleteUserCertificates.ps1","confirm_clearbrowsercache.ps1","DeleteUserCertificates.ps1","AVD_BrowserBookmarkUtility.ps1","AVD_CustomizedStartMenu.ps1")

foreach ($file in $jeaModuleFiles) {
    Copy-Item -Path "$($workdir)\$($file)" -Destination $jeaModulePath -Force | Out-Null

    If (Test-Path -Path "$($jeaModulePath)\$($file)"){
        Write-Output "Successfully copied $($file) into $($jeaModulePath)..."
    }Else{
        Write-Output "Could NOT copy $($file) into $($jeaModulePath)..."
    }
}

<#
#//copy JEA user context scripts into $armyJEAPath
$jeaUserContextScripts = @("confirm_clearteams.ps1","Teams_ClearCache-NewApp.ps1","DeleteUserCache.ps1","confirm_DeleteUserCertificates.ps1","confirm_clearbrowsercache.ps1","DeleteUserCertificates.ps1")

foreach ($file in $jeaUserContextScripts) {
    Copy-Item -Path "$($workdir)\$($file)" -Destination $ArmyJEAPath -Force | Out-Null

    If (Test-Path -Path "$($ArmyJEAPath)\$($file)"){
        Write-Output "Successfully copied $($file) into $($ArmyJEAPath)..."
    }Else{
        Write-Output "Could NOT copy $($file) into $($ArmyJEAPath)..."
    }
}
#>
#//copy LocalAdmin.psrc into $roleCapabilitiesPath
Copy-Item -Path "$($workdir)\LocalAdmin.psrc" -Destination $roleCapabilitiesPath -Force | Out-Null

If (Test-Path -Path "$($roleCapabilitiesPath)\LocalAdmin.psrc"){
    Write-Output "Successfully copied LocalAdmin.psrc into $($roleCapabilitiesPath)..."
}Else{
    Write-Output "Could NOT copy LocalAdmin.psrc into $($roleCapabilitiesPath)..."
}

# Register JEA session but do not restart WinRM
$jeaRegisterPath = "$jeaModulePath\GCCJeaEndpoint.pssc"

Register-PSSessionConfiguration -Name 'GCCJeaEndpoint' -Path $jeaRegisterPath -NoServiceRestart -ErrorAction SilentlyContinue
Write-Output "GCCJeaEndpoint registered. Completed SelfHelpSupplemental files..."
write-output "Completed setup and configuration of JEA module."

Write-Output "Beginning setup and configuration of AUDS Self-Help Tool..."

#//copy AUDS_SelfHelp.exe into $SelfHelpFolderPath
Copy-Item -Path "$($workdir)\AUDS_SelfHelp.exe" -Destination $SelfHelpFolderPath -Force | Out-Null

If (Test-Path -Path "$($SelfHelpFolderPath)\AUDS_SelfHelp.exe"){
    Write-Output "Successfully copied AUDS_SelfHelp.exe into $($SelfHelpFolderPath)..."
}Else{
    Write-Output "Could NOT copy AUDS_SelfHelp.exe into $($SelfHelpFolderPath)..."
}

#//copy shortcut for Self Help in public desktop and start menu;
<#
Write-Output "Creating public desktop and start menu shortcuts for Self Help tool..."
Copy-Item -Path "$($workdir)\SelfHelp.lnk" -Destination "C:\Users\Public\Desktop\SelfHelp.lnk" -Force
Copy-Item -Path "$($workdir)\SelfHelp.lnk" -Destination "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\SelfHelp.lnk" -Force
#>

<#
# Creating Desktop and Start Menu shortcuts
Write-Output "creating Self Help Shortcuts..."
$desktopShortcutPath = "C:\Users\Public\Desktop\SelfHelp.lnk"
$startMenuShortcutPath = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\SelfHelp.lnk"
$SelfHelpExePath = Join-Path -Path $SelfHelpFolderPath -ChildPath "AUDS_SelfHelp.exe"

# Create a shortcut on the Public Desktop
$WScriptShell = New-Object -ComObject WScript.Shell
$shortcut = $WScriptShell.CreateShortcut($desktopShortcutPath)
$shortcut.TargetPath = $SelfHelpExePath
$shortcut.IconLocation = "$SelfHelpExePath,0"
$shortcut.Save()

# Create a start menu shortcut
$startMenuShortcut = $WScriptShell.CreateShortcut($startMenuShortcutPath)
$startMenuShortcut.TargetPath = $SelfHelpExePath
$startMenuShortcut.IconLocation = "$SelfHelpExePath,0"
$startMenuShortcut.Save()
#>

Write-Output "Completed setup and configruation of Self-Help tool"

Stop-Transcript

#//change path back to location before $workdir set;
If($workDir){
    Set-Location -Path $origLocation
}