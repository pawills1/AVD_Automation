$workDir = "C:\temp\bootstrap\DTSLite"
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
    Write-Output "EXE found, create application folder and copying executable there"

    $DTPPath = "$env:ALLUSERSPROFILE\DTPLite"
    $allStartMenu = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs"
    $dskscPath = "$env:Public\Desktop"

    If (!(Test-Path -Path $DTPPath)){
        Write-Output "Creating $($DTPPath) directory"
        New-Item -Path "$($env:ALLUSERSPROFILE)" -Name "DTPLite" -ItemType Directory -Force
    }

    Copy-Item -Path $installFile.FullName -Destination $DTPPath -Verbose -Force -Confirm:$false

    $tgtFile = Get-ChildItem -Path "$DTPPath\$($installfile.Name)"

    If (Test-Path -Path $tgtFile){
        Write-Output "Installation complete, creating Shortcuts"

        $allStartMenu = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs"
        $Target = "$tgtFile"
        $smFolder = "$allStartMenu\DTP Lite"
        $cscPath = "$smFolder\DTP Lite.lnk"

        # Create Start Menu Shortcut
        if (!(Test-Path -LiteralPath $smFolder)){
            Write-Output "Creating Start Menu folder for application shortcut"
            mkdir $smFolder -Verbose -Force
        }
    
        Write-Output "Creating application shortcut"
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.Createshortcut($cscPath)
        $shortcut.TargetPath = $Target
        $shortcut.IconLocation = "$target,0"
        $shortcut.Save()


    }Else{
        Write-Output "Installation failed."
    }

}
else {
    #Assume MSI - standard silent arguments
    Write-Output "Not an EXE, assumed MSI, beginning install with silent switches"
    Start-Process -FilePath $installFile.FullName -ArgumentList "/qn" -Wait
} 

Stop-Transcript

#//change path back to location before $workdir set;
If($workDir){
    Set-Location -Path $origLocation
}
