$workDir = "C:\temp\bootstrap\WinSAPGUI"
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
    $installFile = $InstallerPath | Where-Object { $_.Extension -match ".exe" -and $_.Name -notlike "VC_redist*" }
    #$VCRedisFile = $InstallerPath | Where-Object { $_.Extension -match ".exe" -and $_.Name -like "VC_redist*" }
}
else {
    #The install file is the only file in there.
    $installFile = $InstallerPath
}

#//run command to change user to install mode;
Invoke-Expression -Command "cmd.exe /C change user /install"

#Determine if EXE or MSI - use specific installer switches for each.
if ($installFile -match "exe") {
    <#
    Write-Output "Installing $($VCRedisFile.Name)"
    Start-Process -FilePath $VCRedisFile.FullName -ArgumentList "/install /quiet /norestart" -Wait
    #>
    Write-Output "EXE found, beginning installation with provided switches."
    Start-Process -FilePath $installFile.FullName -ArgumentList "/verysilent /norestart" -Wait
}
else {
    #Assume MSI - standard silent arguments
    Write-Output "Not an EXE, assumed MSI, beginning install with silent switches"
    Start-Process -FilePath $installFile.FullName -ArgumentList "/qn" -Wait
} 

#//run command to change user to install mode;
Invoke-Expression -Command "cmd.exe /C change user /execute"

Stop-Transcript

#//change path back to location before $workdir set;
If($workDir){
    Set-Location -Path $origLocation
}

EXIT 0