$workDir = "C:\temp\bootstrap\IADS4"
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
    # Firefox specific arguments
    Write-Output "EXE found, beginning installation with provided switches."
    Start-Process -FilePath $installFile.FullName -ArgumentList "/S" -Wait
}
else {
    #Assume MSI - standard silent arguments
    Write-Output "Not an EXE, assumed MSI, beginning install with silent switches"
    Start-Process -FilePath $installFile.FullName -ArgumentList "/qn" -Wait
} 

#//Setting suggested IADS4 workspace location in registry...
$wkspcefilePath = "$env:ProgramData\IADS4_Workspace"
$defaultWkspce = "$env:SystemDrive\IADS4_Workspace"
$wkspcereg = @{
    Path  = "HKLM\SOFTWARE\IADS"
    Type  = "REG_SZ"
    Name  = "Workspace"
    Value = "$wkspcefilePath" 
}

write-output "Setting suggested IADS4 workspace location in registry..."
REG ADD $wkspcereg.Path /v $wkspcereg.Name /t $wkspcereg.Type /d $wkspcereg.Value /reg:64 /f

#//removing default workspace location created by IADS4 install...
if (Test-Path -Path $defaultWkspce) {
    write-output "removing default workspace location created by IADS4 install..."
    try {
        Write-Output "Removing default workspace."
        Remove-Item -Path $defaultWkspce -Force -Verbose -ErrorAction Stop
    }
    catch {
        if ($error[0]) {
            Write-Output "Here's what went wrong, but it can be safely ignored:"
            $_
        }
    }
}

#//creating suggested workspace location for IADS4 install...
if (-Not (Test-Path -Path $wkspcefilePath)){
    write-output "creating suggested workspace location for IADS4 install..."
    try{
        mkdir $wkspcefilePath -Verbose -Force -Confirm:$false -ErrorAction Stop
    }catch{
        if ($error[0]){
            Write-Output "Unable to create $wkspcefilePath. See error:$_"
        }
    }
}

Stop-Transcript

#//change path back to location before $workdir set;
If($workDir){
    Set-Location -Path $origLocation
}
