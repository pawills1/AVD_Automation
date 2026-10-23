$workDir = "C:\temp\bootstrap\FedLog"
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

#//create FEDLOG program directory, if not there
If (!(test-path -Path "C:\ProgramData\FEDLOG")){
    new-item -Path "C:\ProgramData\FEDLOG" -ItemType Directory -Force
}

$InstallerPath = Get-childitem -path . -File

If($useZIPfile -eq $TRUE){
    Write-Output "Extracting contents of $($ZIPFile) to C:\ProgramData\FEDLOG..."
    $ZIPFile = $InstallerPath | Where-Object {$_.Extension -match ".zip"}
    Expand-Archive -literalPath $ZIPFile -DestinationPath "C:\ProgramData\FEDLOG" -Force
}

#//creating application shortcut;
New-Item -Type SymbolicLink -Path "C:\Users\Public\Desktop\FED LOG" -Target "C:\ProgramData\FEDLOG\IMD64.EXE" -Force

Stop-Transcript

#//change path back to location before $workdir set;
If($workDir){
    Set-Location -Path $origLocation
}
