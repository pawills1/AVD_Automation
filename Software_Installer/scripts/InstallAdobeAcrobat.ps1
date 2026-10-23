$workDir = "C:\temp\bootstrap\AdobeAcrobat"
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
    #$ZIPFile = "$($workDir)\AdobeAcrobat_AVD.zip"
    Expand-Archive -literalPath $ZIPFile -DestinationPath $workDir -Force
}

$InstallerPath = Get-childitem -path . -File

#//installers;
$installFile = "`"$($workDir)\setup.exe`""
$licenseEXE  = "`"$($workDir)\adobe-licensing-toolkit.exe`""
$JSONFile    = "`"$($workDir)\ngl-preconditioning-data.json`""
$reg         = "`"$($workDir)\AdobeSettings.reg`""
$patchFile = $InstallerPath | Where-Object {$_.Extension -match ".msp"}

#//application installed, continue with additional customizations;
Write-Output "EXE found, beginning installation with provided switches."
Start-process $installFile -ArgumentList "/sALL /msi /qn /l*v `"$($logFolder)\$($appNameForLog)_AVD_AppInstall.log`"" -Wait
    
Write-Output "installing update $($patchfile.Name)..."
Start-Process $patchFile.FullName -ArgumentList "/qn /l*v `"$($logFolder)\$($patchfile.BaseName).log`"" -Wait

Write-Output "applying Adobe Acrobat licensing..."
Start-Process $licenseEXE -ArgumentList "--precondition --install  --filePath $JSONFile" -wait

Write-Output "applying Adobe Acrobat settings via .reg files..."
#Start-Process -filepath "$env:windir\regedit.exe" -Argumentlist @("/s", "$reg") -Verbose -Wait

Stop-Transcript

#//change path back to location before $workdir set;
If($workDir){
    Set-Location -Path $origLocation
}

