$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

$workDir = "C:\temp\bootstrap\AppxStateRepositoryDBRebuild"
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

$InstallerPath = Get-childitem -path . -File

If($useZIPfile -eq $TRUE){
    $ZIPFile = $InstallerPath | Where-Object {$_.Extension -match ".zip"}
    Write-Output "Extracting contents of $($ZIPFile) to $($workDir)..."
    Expand-Archive -literalPath $ZIPFile -DestinationPath $workDir -Force
}

$PSExec = "$($workDir)\PsExec64.exe"
$script = "$($workDir)\RunStateRepositoryDBBackup.ps1"

$arguments = @(
    "-acceptEULA",
    "-s",
    "-i",
    "Powershell.exe",
    "-ExecutionPolicy Bypass -File `"$($script)`""
)

Start-Process $PSExec -ArgumentList $arguments -Wait