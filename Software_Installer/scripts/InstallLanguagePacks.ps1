Function get-timestamp{

    return get-date -Format "MM/dd/yy HH:mm:ss"

}#//end function

$langPackList = "ko-kr;ja-jp;fr-fr;de-de"
$useZipFile   = $TRUE
$workDir      = "C:\temp\bootstrap\LanguagePacks"
$regKey       = "HKLM:\SOFTWARE\ArmyAVD"

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

If(!(test-path -Path $regKey)){
    New-Item -Path $regKey -Force | Out-Null
}

#//do until loop to find next language pack out of $arrLangPacks not installed;
$arrLangPacks = $langPackList.split(';')

ForEach($language in $arrLangPacks){
    
    $regVal = Get-ItemProperty -Path $regKey -name $language -ErrorAction silentlycontinue

    If(!($regVal)){
        break
    }

}

$logFolder = "$env:SystemDrive\Program Files\AVDAppInstallLogs"
$appNameForLog = $($workDir).split('\')[3]
$LogPath = "$($logFolder)\$($appNameForLog)_$($language)_AVD_Script.log"

Start-Transcript -Path $logPath

Write-Output "$($language) not installed, installing $($language) LP and basic typing capability..."

$InstallerPath = Get-childitem -path . -File

If($useZIPfile -eq $TRUE){
    Write-Output "Extracting contents of $($ZIPFile) to $($workDir)..."
    $ZIPFile = $InstallerPath | Where-Object {$_.Extension -match ".zip"}
    Expand-Archive -literalPath $ZIPFile -DestinationPath $workDir -Force
}

#//Disable Language Pack Cleanup;
Write-Output "disabling pre-staged app cleanup scheduled tasks..."
Disable-ScheduledTask -TaskPath "\Microsoft\Windows\AppxDeploymentClient\" -TaskName "Pre-staged app cleanup"
Disable-ScheduledTask -TaskPath "\Microsoft\Windows\MUI\" -TaskName "LPRemove"
Disable-ScheduledTask -TaskPath "\Microsoft\Windows\LanguageComponentsInstaller" -TaskName "Uninstallation"
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Control Panel\International" /v "BlockCleanupOfUnusedPreinstalledLangPacks" /t REG_DWORD /d 1 /f
            
#//install language pack;
$LP = gci -Path $workdir | ? {($_.name -like '*Language-Pack*') -AND ($_.name -like "*$($language)*")}

Write-Output "$(get-timestamp): Installing Language Pack $($LP.name)..."
Dism /Online /Add-Package /PackagePath:$($LP.FullName)  | Out-Null

#//add language capability features;
$langFeatures = gci -Path $workdir | ? {($_.name -like '*LanguageFeatures*') -AND ($_.name -like "*$($language)*") -AND ($_.name -notlike '*Fonts*')}

ForEach($package in $langFeatures){
                
    switch -Wildcard ($package){
                    
        "*Basic*"{
            Write-Output "$(get-timestamp): adding language capability $($package.name)..."
            Dism /Online /Add-Capability /CapabilityName:"Language.Basic~~~$($language)~0.0.1.0" /Source:$($package.FullName) | Out-Null
        }

    }

}

#//confirm langauge support has been successfully added and write info to registry;
$langInfo = dism /online /get-intl /English | Select-String -SimpleMatch 'Installed language(s)'
If($langInfo -match "$($language)"){
    $regVal = get-timestamp
    New-ItemProperty -Path $regKey -Name $language -Value $regVal
}

Stop-Transcript

#//change path back to location before $workdir set;
If($workDir){
    Set-Location -Path $origLocation
}

