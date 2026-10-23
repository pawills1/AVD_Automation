$workDir    = "C:\temp\bootstrap\M365CoPilot"
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

$InstallerPath = Get-childitem -path . -File

If($installerPath.count -gt 1){
    $installFile = $installerPath | Where-Object {$_.Extension -match ".appxbundle"}
    $dependencies = Get-ChildItem "$($workDir)\Dependencies" | Select-Object -ExpandProperty FullName
    Add-AppxProvisionedPackage -Online -PackagePath "$($installFile.FullName)" -DependencyPackagePath $dependencies -SkipLicense
}

Stop-Transcript

#//change path back to location before $workdir set;
If($workDir){
    Set-Location -Path $origLocation
}