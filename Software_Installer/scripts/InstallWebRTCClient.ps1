$workDir = "C:\temp\bootstrap\WebRTCClient"
$useZIPfile = $false

#//if $workdir specified, sets that as working location (if it exists). Otherwise, assumes script location is working directory.
If($workDir){

    If(Test-Path -Path "$($workdir)"){
        $origLocation = (Get-Location).path
        Set-Location -Path $($workDir)
    }else{
        If ($useZIPfile -eq $false){
            Write-Output "The working path $($workDir) does not exit. Creating it now"
            New-Item -ItemType Directory -Path $workDir -Force
            Set-Location -Path $($workDir)
        }else{
            Write-Warning "The working path $($workDir) does not exit. Script will exit."
            exit
        }
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
}else{
    #//$useZIPFile set to $false, download installers from internet;
    Invoke-WebRequest -Uri "https://aka.ms/msrdcwebrtcsvc/msi" -OutFile "$($workdir)\WebRTCSetup.msi"
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
    # Firefox specific arguments
    Write-Output "EXE found, beginning installation with provided switches."
    Start-Process -FilePath $installFile.FullName -ArgumentList "/S" -Wait
}
else {
    #Assume MSI - standard silent arguments
    Write-Output "Not an EXE, assumed MSI, beginning install with silent switches"
    Start-Process -FilePath "C:\Windows\System32\msiexec.exe" -ArgumentList "/I $($installFile.FullName) /qn /norestart /l*v `"$($logFolder)\$($appNameForLog)_AppInstall.log`"" -Wait
} 

Stop-Transcript

#//change path back to location before $workdir set;
If($workDir){
    Set-Location -Path $origLocation
}
