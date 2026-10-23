Function get-timestamp{

    return get-date -Format "MM/dd/yy HH:mm:ss"

}#//end function

$useZipFile   = $TRUE
$workDir      = "C:\temp\bootstrap\LangPackJA"

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

#//get lang from $appNameForLog variable;

$langFromFolder = $appNameForLog.Substring($appNameForLog.Length - 2).ToLower()
switch($langFromFolder){
    
    "ko"{$language = "$($langFromFolder)-kr"};

    "ja"{$language = "$($langFromFolder)-jp"};

    "fr"{$language = "$($langFromFolder)-fr"};

    "de"{$language = "$($langFromFolder)-de"};

}

$InstallerPath = Get-childitem -path . -File

If($useZIPfile -eq $TRUE){
    $ZIPFile = $InstallerPath | Where-Object {$_.Extension -match ".zip"}
    Write-Output "Extracting contents of $($ZIPFile) to $($workDir)..."
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
Dism /Online /Add-Package /PackagePath:$($LP.FullName)  #| Out-Null

#//add language capability features;
$langFeatures = gci -Path $workdir | ? {($_.name -like '*LanguageFeatures*') -AND ($_.name -like "*$($language)*") -AND ($_.name -notlike '*Fonts*')}

ForEach($package in $langFeatures){
                
    switch -Wildcard ($package){
                    
        "*Basic*"{
            Write-Output "$(get-timestamp): adding language capability $($package.name)..."
            Dism /Online /Add-Capability /CapabilityName:"Language.Basic~~~$($language)~0.0.1.0" /Source:$($package.FullName) #| Out-Null
        }

    }

}

Stop-Transcript

#//change path back to location before $workdir set;
If($workDir){
    Set-Location -Path $origLocation
}

EXIT 0
