$workDir = "C:\temp\bootstrap\Firefox"
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

#determine installation log path;
If(!(test-path -Path "$env:SystemDrive\Program Files\AVDAppInstallLogs")){
    New-Item -ItemType Directory -Path "$env:SystemDrive\Program Files\AVDAppInstallLogs" -Force
}

$logFolder = "$env:SystemDrive\Program Files\AVDAppInstallLogs"
$appNameForLog = $($workDir).split('\')[3]
$LogPath = "$($logFolder)\$($appNameForLog)_AVD_Script.log"

Start-Transcript -Path $logPath

$InstallerPath = Get-childitem -path . -File

If($useZIPfile -eq $TRUE){
    Write-Output "Extracting contents of $($ZIPFile) to $($workDir)..."
    $ZIPFile = $InstallerPath | Where-Object {$_.Extension -match ".zip"}
    Expand-Archive -literalPath $ZIPFile -DestinationPath $workDir -Force
}

$InstallerPath = Get-childitem -path . -File
    
If ($InstallerPath.count -gt 1) {
    #If more than one file found, filter for installer.
    Write-Output "More than one file, filtering for the installer."
    $installFile = $InstallerPath | Where-Object { $_.Extension -match ".exe" -or $_.Extension -match ".msi" }
}
else {
    #The install file is the only file in there.
    $installFile = $InstallerPath
}


#Determine if EXE or MSI - use specific installer switches for each.
if ($installFile -match "exe") {
    Write-Output "EXE found, beginning installation with provided switches."
    Start-Process -FilePath $installFile -ArgumentList "/s" -Wait
}
else {
    #Assume MSI - standard silent arguments
    Write-Output "Not an EXE, assumed MSI, beginning install with silent switches"
    $appLog = "`"$($logFolder)\$($appNameForLog)_AVD_AppInstall.log`""
    Start-Process -FilePath "C:\Windows\System32\msiexec.exe" -ArgumentList "/I $($installFile) PRIVATE_BROWSING_SHORTCUT=false INSTALL_MAINTENANCE_SERVICE=false /q /norestart /L*V $($appLog)" -Wait
    
    <#
    $reg = $InstallerPath | Where-Object {$_.Extension -match ".reg"}
    If ($reg){
        Write-Output "Reg file found, installing registry configuration."
        Start-Process -filepath "$env:windir\regedit.exe" -Argumentlist @("/s", "$reg") -Verbose
    }else{
        #Write to the log that the reg file could not be found.
        Write-Output "Registry file not found."
    }
    
    $ext = $InstallerPath | Where-Object {$_.Extension -match ".xpi"}
    If($ext){
        Write-Output "Copying extension into Firefox features directory"
        Copy-Item -Path $ext.fullName -Destination "$env:ProgramFiles\Mozilla Firefox\browser\features" -Force
    }
    #>

    #//remove Mozilla Updater tool, if setting propety in MSI didn't remove it;
    If (Test-Path -Path "C:\Program Files (x86)\Mozilla Maintenance Service\Uninstall.exe"){
        Start-Process -FilePath "C:\Program Files (x86)\Mozilla Maintenance Service\Uninstall.exe" -ArgumentList "/S"
    }

    Write-Output "Creating Firefox cookie exception for Menlo Security page"
    $cookies = "HKLM\Software\Policies\Mozilla\Firefox\Cookies"
    $cookiesAllow = "$cookies\Allow"
    $cookieallowSession = "$cookies\AllowSession"
    $cookieBlock = "$cookies\Block"

    REG add $cookies /reg:64 /f
    REG add $cookies /v "Behavior" /t REG_SZ /d "reject-foreign" /reg:64 /f
    REG add $cookies /v "Locked" /t REG_DWORD /d 1 /reg:64 /f
    REG add $cookieallowSession /v "1" /t REG_SZ "" /reg:64 /f
    REG add $cookieBlock /v "1" /t REG_SZ "" /reg:64 /f
    REG add $cookiesAllow /v "1" /t REG_SZ /d "https://menlosecurity.com" /reg:64 /f
    REG add $cookiesAllow /v "2" /t REG_SZ /d "http://menlosecurity.com" /reg:64 /f

}

Stop-Transcript

#//change path back to location before $workdir set;
If($workDir){
    Set-Location -Path $origLocation
}