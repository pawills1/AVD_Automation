$destinationPath = "C:\Program Files\Geekly"
$useZIPfile = $true
$rdAgentDirectories = @("C:\temp\bootstrap\RDAgentBootloader",
                        "C:\temp\bootstrap\RDAgentInstaller")

if (!(Test-Path -Path $destinationPath)){
    New-Item -ItemType Directory -Path $destinationPath -Force
}

#determine installation log path;
if (!(Test-Path -Path "$env:SystemDrive\Program Files\AVDAppInstallLogs")){
    New-Item -ItemType Directory -Path "$env:SystemDrive\Program Files\AVDAppInstallLogs" -Force
}

$logFolder = "$env:SystemDrive\Program Files\AVDAppInstallLogs"

#//if $workdir specified, sets that as working location (if it exists). Otherwise, assumes script location is working directory.
foreach ($dir in $rdAgentDirectories){
    $appNameForLog = $($dir).split('\')[3]
    $logPath = "$($logFolder)\$($appNameForLog)_AVD_Script.log"
    Start-Transcript -Path $logPath

    if ($dir){
        if (Test-Path -Path "$($dir)"){
            $originalLocation = (Get-Location).Path
            Set-Location -Path $($dir)
        }else{
            Write-Warning "The working path '$($dir)' does not exit. Script will exit."
            exit
        }
    }

    $zipPath = Get-ChildItem -Path . -File
    if ($useZIPfile){
        Write-Output "Extracting contents of '$($zipPath.Name)' to '$($destinationPath)'."
        $zipFile = $zipPath | Where-Object {$PSItem.Extension -match ".zip"}
        Expand-Archive -literalPath $zipFile -DestinationPath $destinationPath -Force
    }

    $msi = Get-ChildItem -Path $destinationPath -Filter "*.msi"
    if (!$msi){
        Write-Output "MSI for '$($dir.split("\")[-1])' was not found."
        break
    }
    Write-Output "MSI '$($msi[0].Name)' was extracted from '$($dir)' and exists in '$($destinationPath)'."
    Stop-Transcript
}

#//change path back to location before $workdir set;
if($workDir){
    Set-Location -Path $originalLocation
}
