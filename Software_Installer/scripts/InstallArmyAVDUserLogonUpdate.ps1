$workDir = "C:\temp\bootstrap\ArmyAVDUserLogonUpdate"
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

Write-Output "copying updated ArmyAVDUserLogon.exe to C:\Program Files\ArmyAVD"
copy-item "$($workDir)\ArmyAVDUserLogon.exe" -destination "C:\Program Files\ArmyAVD" -Force

Write-Output "copying OUtlook OWA SMIME control folder for Edge and Chrome to C:\Program Files\ArmyAVD"
Copy-Item "$($workDir)\SmimeOutlookWebChrome.msi" -Destination "C:\Program Files\ArmyAVD" -Recurse -Force

Write-Output "copying AVDAppLockerRules.xml to C:\Program Files\ArmyAVD"
Copy-Item "$($workDir)\AVDApplockerRules.xml" -Destination "C:\Program Files\ArmyAVD" -Force


#//add S/MIME extension for Chrome;
Write-Output "Adding SMIME OWA extension for Chrome"

$extensionID = "maafgiompdekodanheihhgilkjchcakm"
$updateURL   = "https://outlook.office.com/owa/SmimeCrxUpdate.ashx"

$regPath = "HKLM:\SOFTWARE\Policies\Google\Chrome\ExtensionInstallForcelist"

If(!(Test-Path -Path $regPath)){
    New-Item -Path $regPath -force
}
$regVal = "$extensionID;$updateURL"
Set-ItemProperty -Path $regPath -Name "2" -Value $regVal


#//add disableMSI setting so SMIME OWA control can run in user context;
$regPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Installer"

If(!(Test-Path -Path $regPath)){
    New-Item -Path $regPath -force
}
$regVal = 0
Set-ItemProperty -Path $regPath -Name "DisableMSI" -Value $regVal

#//add new applocker policy that includes publsher rule for SMIME OWA conrol and assocaited files;
Write-Output "applying updated applocker rules to support SMIME through OWA"
$appLockerXML  = "C:\Program Files\ArmyAVD\AVDApplockerRules.xml"
Set-AppLockerPolicy -XMLPolicy $($appLockerXML)


If ($doLogging -eq $TRUE){
    start-transcript -path $logPath
}
