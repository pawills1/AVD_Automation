$workDir = "C:\temp\bootstrap\MicrosoftStoreAppUpdates"
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
    $ZIPFile = $InstallerPath | Where-Object {$_.Extension -match ".zip"}
    Write-Output "Extracting contents of $($ZIPFile) to $($workDir)..."
    Expand-Archive -literalPath $ZIPFile -DestinationPath $workDir -Force
}

$InstallerPath = Get-childitem -path . -File

#$arrStoreAppsToUpdate = @("Microsoft.Windows.Photos","Microsoft.ScreenSketch")
$arrStoreAppsToUpdate = @("Microsoft.Windows.Photos")

ForEach ($storeApp in $arrStoreAppsToUpdate){
    
    $installFile = $InstallerPath | Where-Object {$_.Name -match "$($storeApp)"}

    #//create depenecies list for app based on app name;
    switch($storeApp){
        
        "Microsoft.Windows.Photos"{
            $dependencyList = $InstallerPath | Where-Object {($_.name -like "*UWPDesktop*") -OR ($_.name -like "*VCLibs.140.00_14.0.33519.0*") -OR ($_.name -like "*WindowsAppRuntime*")}
        };

        "Microsoft.ScreenSketch"{
            $dependencyList = $InstallerPath | Where-Object {($_.name -like "*UWPDesktop*") -OR ($_.name -like "*VCLibs.140.00_14.0.33519.0*") -OR ($_.name -like "*WindowsAppRuntime*") -OR ($_.name -like "*UI.Xaml*")}
        }

    }

    Write-Output "Deprovisioning existing version of $($storeApp) for system..."
    try{
        Get-AppxProvisionedPackage -Online | ? {$_.DisplayName -eq "$($storeApp)"} | Remove-AppxProvisionedPackage -Online
        Write-Output "successfully deprovisioned $($storeApp) app from system..."
    }catch{
        Write-Warning "Could not successfully deprovision $($storeApp) from system."
    }

    Write-Output "Uninstalling existing version of $($storeApp) for all users on system..."
    try{
        Get-AppxPackage -AllUsers | ? {$_.Name -eq  "$($storeApp)"} | Remove-AppxPackage -AllUsers
        Write-Output "successfully uinstalled $($storeApp) app"
    }catch{
        Write-Warning "Could not successfully uninstall $($storeApp)"
    }

    Write-Output "Provisioning $($installFile.name) on system..."
    Try{
        Add-AppxProvisionedPackage -Online -PackagePath $installFile.FullName -DependencyPackagePath $dependencyList.FullName -SkipLicense
        Write-Output "Successfully provisioned Microsoft.Windows.photos and dependencies to system"
    }catch{
        Write-Warning "Could not successfully provision $($installFile.name) and/or dependencies to system"
        Write-Output $_.Exception.Message
    }

}

Stop-Transcript