$workDir = "C:\temp\bootstrap\ClassificationBanner"
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
    $ZIPFile = $InstallerPath | Where-Object {$_.Extension -match ".zip"}
    Write-Output "Extracting contents of $($ZIPFile) to $($workDir)..."
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
    # Change AVD to install mode
    $appLog = "`"$($logFolder)\$($appNameForLog)_AVD_AppInstall.log`""
    Start-Process -FilePath "C:\Windows\System32\msiexec.exe" -ArgumentList "/I $($installFile) /qn BANNERRED=0 BANNERGREEN=122 BANNERBLUE=51 BANNERTEXT=`"CONTROLLED UNCLASSIFIED INFORMATION`" BOUNCY=1 BOUNCEBACK=1 FONTRED=255 FONTGREEN=255 FONTBLUE=255 BOLD=1 UNDERLINE=0 ITALICS=0 STRIKEOUT=0 BANNERPADDING=4 /L*V $($appLog)" -Wait

}

Stop-Transcript
    
#//change path back to location before $workdir set;
If($workDir){
    Set-Location -Path $origLocation
}