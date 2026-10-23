Function test-regval{

    param(
        [Parameter(Mandatory=$True)]
        [string]$regPath,
        [Parameter(Mandatory=$True)]
        [string]$regSetting,
        [Parameter(Mandatory=$True)]
        [string]$regVal
    )

    $chkRegVal = Get-ItemProperty -Path $RegPath -Name $regSetting -ErrorAction SilentlyContinue | Select-Object -ExpandProperty $regSetting

     If ($chkRegVal -eq $regVal){
        return $true
     }Else{
        return $False
     }
     
} #//end function


###################################################################################################

$workDir = "C:\temp\bootstrap\LGPOSettings"
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

#get list of unzipped contents;
$InstallerPath = Get-childitem -path . -File

#//create C:\Program Files\ArmyAVD directory if not there;
Write-Output "creating C:\Program Files\ArmyAVD"
If (!(Test-Path -Path "C:\Program Files\ArmyAVD")){
    mkdir "C:\Program Files\ArmyAVD" -ErrorAction SilentlyContinue
}

#//copy items into C:\Program Files\ArmyAVD, theme & wallpaper directores;
<#
write-output "Copying files into C:\Program Files\ArmyAVD and C:\Windows\Web\Wallpaper\Windows"
copy-item "$($workdir)\AVDUserSettingsRegistry.pol" -destination "C:\Program Files\ArmyAVD" -Force
copy-item "$($workdir)\AVDComputerSecuritySettings.inf" -destination "C:\Program Files\ArmyAVD" -Force
copy-item "$($workdir)\AVDAppLockerRulesV2.xml" -destination "C:\Program Files\ArmyAVD" -Force
copy-item "$($workdir)\ArmyAVDUserLogon.ps1" -destination "C:\Program Files\ArmyAVD" -Force
copy-item "$($workdir)\LGPO.exe" -destination "C:\Program Files\ArmyAVD" -Force
copy-item "$($workdir)\Army.png" -destination "C:\Windows\Web\Wallpaper\Windows" -Force
copy-item "$($workdir)\SmimeOutlookWebChrome.msi" -destination "C:\Program Files\ArmyAVD" -Force
copy-item "$($workdir)\MSTeams.ico" -destination "C:\Program Files\ArmyAVD" -Force
copy-item "$($workdir)\logout-icon-1.ico" -destination "C:\Program Files\ArmyAVD" -Force
copy-item "$($workdir)\logout-icon-2.ico" -destination "C:\Program Files\ArmyAVD" -Force
copy-item "$($workdir)\logout-icon-4.ico" -destination "C:\Program Files\ArmyAVD" -Force
copy-item "$($workdir)\Outlook.ico" -destination "C:\Program Files\ArmyAVD" -Force
copy-item "$($workdir)\Edge.ico" -destination "C:\Program Files\ArmyAVD" -Force
copy-item "$($workdir)\armyavdusersettings.ADMX" -destination "C:\Windows\PolicyDefinitions" -Force
copy-item "$($workdir)\armyavdusersettings.ADML" -destination "C:\Windows\PolicyDefinitions\en-US" -Force
copy-item "$($workdir)\Local State" -destination "C:\Program Files\ArmyAVD" -Force
copy-item "$($workdir)\UAT_URLs.txt" -destination "C:\Program Files\ArmyAVD" -Force
copy-item "$($workdir)\GenAI.ico" -destination "C:\Program Files\ArmyAVD" -Force
copy-item "$($workdir)\GenAI_Popup_Prompt.ps1" -destination "C:\Program Files\ArmyAVD" -Force
copy-item "$($workdir)\GenAI_Popup_Picture.png" -destination "C:\Program Files\ArmyAVD" -Force
copy-item "$($workdir)\Copilot.ico" -destination "C:\Program Files\ArmyAVD" -Force
#>
gci -Path $workDir -File | ? {($_.name -ne "LGPOSettings_AVD.zip") -AND ($_.name -notlike "armyavdusersettings*") -AND ($_.name -ne "Army.png")} | Copy-Item -destination "C:\Program files\ArmyAVD" -Force
copy-item "$($workdir)\armyavdusersettings.ADMX" -destination "C:\Windows\PolicyDefinitions" -Force
copy-item "$($workdir)\armyavdusersettings.ADML" -destination "C:\Windows\PolicyDefinitions\en-US" -Force
copy-item "$($workdir)\Army.png" -destination "C:\Windows\Web\Wallpaper\Windows" -Force

#//create Run key value to ensure AVDUserLogon.exe runs for each user logging into VM;
write-output "Writing ArmyAVDUserLogon and GenAIUserPrompt into Run key..."
Set-ItemProperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -Name "ArmyAVDUserLogon" -Value "cmd /c start /min `"`" powershell -WindowStyle Hidden -ExecutionPolicy Bypass -NoProfile -File `"C:\Program Files\ArmyAVD\ArmyAVDUserLogon.ps1`""
Set-ItemProperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -Name "ArmyAVDGenAIPrompt" -Value "cmd /c start /min `"`" powershell -WindowStyle Hidden -ExecutionPolicy Bypass -NoProfile -File `"C:\Program Files\ArmyAVD\GenAI_Popup_Prompt.ps1`" -Timeout 0 -ImagePath `"C:\Program Files\ArmyAVD\GenAI_Popup_Picture.png`" " -Force

Write-Output "Running LGPO to apply system and user GPO settings"
#//info for LGPO;
$regPOLFile  = "$($workdir)\AVDUserSettingsRegistry.pol"
$GptTmplFile = "$($workdir)\AVDComputerSecuritySettings.inf"
$LGPOFile    = "C:\Program Files\ArmyAVD\LGPO.exe"

#//apply user based regsitry .pol file settings using LGPO;
Start-Process -FilePath $($LGPOFile) -ArgumentList "/u $($regPOLFile)" -Wait

#//apply computer based regsitry .pol file settings using LGPO;
Start-Process -FilePath $($LGPOFile) -ArgumentList "/s $($GptTmplFile)" -Wait

#//set AppIDsvc to auto, and start it. Dependency for applocker;
Write-Output "settings AppIDSvc to automatic startup type, dependency for AppLocker..."
$cmd1 = "sc.exe config AppIDSvc start= auto"
Invoke-expression -command $cmd1
Set-Service -Name AppIDSvc -Status Running -PassThru -ErrorAction SilentlyContinue | Out-Null

write-output "Setting hidden folder attribute on C:\Program Files\ArmyAVD..."
&attrib.exe +h +s "C:\Program Files\ArmyAVD"

#//apply AVD related applocker rules;
Write-Output "Applying Army AVD AppLocker rules"
$appLockerXML  = "$($workdir)\AVDApplockerRulesV2.xml"
Set-AppLockerPolicy -XMLPolicy $($appLockerXML)

Stop-Transcript

#//change path back to location before $workdir set;
If($workDir){
    Set-Location -Path $origLocation
}