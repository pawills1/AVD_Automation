$install = $true
$workDir = "C:\temp\bootstrap\GoogleChrome"
$useZIPfile = $True

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
    Start-Process -FilePath "C:\Windows\System32\msiexec.exe" -ArgumentList "/I $($installFile) /q /norestart /L*V $($appLog)" -Wait

    Write-Output "Stopping & Disabling Google Updater services..."
    $googleServices = Get-Service | ? {$_.name -like "Google*"}

    ForEach($service in $googleServices){

        if (!($service.Status -eq "Stopped")){
            $cmd = "sc.exe stop $($service.name)"
            Invoke-Expression -Command $cmd -ErrorAction SilentlyContinue
        }
        If(!($service.StartType -eq "Disabled")){
            $cmd = "sc.exe config $($service.name) start=disabled"
            Invoke-Expression -Command $cmd -ErrorAction SilentlyContinue
        }
      
    }
    <#
    $reg = $InstallerPath | Where-Object {$_.Extension -match ".reg"}
    If ($reg){
        Write-Output "Reg file found, installing registry configuration."
        Start-Process -filepath "$env:windir\regedit.exe" -Argumentlist @("/s", "$reg") -Verbose
    }else{
        #Write to the log that the reg file could not be found.
        Write-Output "Registry file not found."
    }
    #>
    
    #//disable option to create chrome desktop shortcut in user profile;
    $InstallLocation = Split-Path ((Get-ItemProperty -Path "HKLM:\SOFTWARE\Classes\ChromeHTML\shell\open\command" -Name "(Default)")."(Default)").Split("--")[0].Trim('"',' ') -Parent
    $PrefFile = "$($InstallLocation)\initial_preferences"

    if (Test-Path -Path "$PrefFile"){
        $json = Get-Content $PrefFile | ConvertFrom-Json
    }

    If ($json.distribution -ne $null) {
        $json.distribution | Add-Member -Name "do_not_create_desktop_shortcut" -Value $true -MemberType NoteProperty -Force
    }Else{
        $json | Add-Member -Name "distribution" -Value @{"do_not_create_desktop_shortcut"=$true} -MemberType NoteProperty -Force
    }
    
    $json | ConvertTo-Json -Depth 100 | Out-File -FilePath $PrefFile -Encoding ASCII

}

Stop-Transcript
    
#//change path back to location before $workdir set;
If($workDir){
    Set-Location -Path $origLocation
}