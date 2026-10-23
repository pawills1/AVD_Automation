$workDir = "C:\temp\bootstrap\OCETool"
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
    Write-Output "EXE found, create C:\OCETOOL folder and copy executable there"
        
    If (!(Test-Path -Path "C:\OCETOOL")){
        Write-Output "Creating C:\OCETOOL directory"
        New-Item -Path "C:\\" -Name "OCETOOL" -ItemType Directory -Force
    }

    Write-Output "Copying OCE.EXE to C:\OCETOOL folder"
    Copy-Item -LiteralPath $installFile.FullName -Destination "C:\OCETOOL" -Force

}
else {
    #Assume MSI - standard silent arguments
    Write-Output "Not an EXE, assumed MSI, beginning install with silent switches"
    Start-Process -FilePath $installFile.FullName -ArgumentList "/qn" -Wait
} 

#//creating start menu and public desktop shortcuts;
<#
Copy-Item -Path "$($workDir)\OCE Tool.lnk" -Destination "C:\Users\Public\Desktop\OCE Tool.lnk" -Force
Copy-Item -Path "$($workDir)\OCE Tool.lnk" -Destination "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OCE Tool.lnk" -Force
#>

<#
Write-Output "creating start menu and public desktop shortcuts..."
$shell = New-Object -ComObject wscript.shell 
$SMshortCut = $shell.CreateShortcut($OCEMSShortcut)
$PBshortcut = $shell.CreateShortcut($OCEPublicShortcut)
$SMShortcut.targetpath = $OCETarget 
$SMShortcut.iconlocation = "$OCETarget,0"
$SMShortcut.Save()
$PBShortcut.targetpath = $OCETarget
$PBShortcut.iconlocation = "$OCETarget,0"
$PBShortcut.Save()
#>

Stop-Transcript

#//change path back to location before $workdir set;
If($workDir){
    Set-Location -Path $origLocation
}
