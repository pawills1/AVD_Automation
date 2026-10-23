$workDir = "C:\temp\bootstrap\Microsoft.Windows.Photos"
$useZIPfile = $TRUE

#//if $workdir specified, sets that as working location (if it exists). Otherwise, assumes script location is working directory.
If($workDir){

    If(Test-Path -Path "$($workdir)"){
        $origLocation = (Get-Location).path
        Set-Location -Path $($workDir)
    }else{
        If ($useZIPfile -eq $false){
            Write-Output "The working path $($workDir) does not exit. Creating it now"
            New-Item -ItemType Directory -Path $workDir -Force
            Set-Location -Path $($workDir)
        }else{
            Write-Warning "The working path $($workDir) does not exit. Script will exit."
            exit
        }
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
}else{
    #//$useZIPFile set to $false, download installers from internet;
    $appInstaller = "Microsoft.Windows.Photos_2024.11100.17007.0_neutral_~_8wekyb3d8bbwe.msixbundle"
    $appURL  = "http://tlu.dl.delivery.mp.microsoft.com/filestreamingservice/files/2e70acc3-2067-4a38-98ff-15aca99d5394?P1=1731091778&P2=404&P3=2&P4=VLvc9BM8R7Z3w9OpcZib03p0T0KCZo5JPl7iCaWQ1IcMJpq6bOr3qfbWFaNekS7AM%2fMd2apDnZU4vydFK%2bchOA%3d%3d"
    Invoke-WebRequest -Uri $appURL -OutFile "$($workdir)\$($appInstaller)"
}
   
$InstallerPath = Get-childitem -path . -File

If ($InstallerPath.count -gt 1) {
    #If more than one file found, filter for installer.
    Write-Output "More than one file, filtering for the installer."
    $installFile = $InstallerPath | Where-Object { $_.Extension -match ".exe" -or $_.Extension -match ".msi" -or $_.Extension -match ".msixbundle" }
}
else {
    #The install file is the only file in there.
    $installFile = $InstallerPath
}

#Determine if EXE or MSI - use specific installer switches for each.
if ($installFile -match "msixbundle") {
    # Firefox specific arguments
    Write-Output "msixbundlefound found, beginning process to removing existing app and install new app..."

    $appInstaller = $installFile.Name
    Write-Output "Uninstalling $($appInstaller.split('_')[0]) for all users..."
    Get-appxpackage -allUsers -name $($appInstaller.split('_')[0]) | remove-appxpackage

    write-output "Deprovisioning $($appInstaller.split('_')[0]) from system..."
    $photosProvisionedApp = get-appxprovisionedPackage -online | ? {$_.displayName -eq "$($appInstaller.split('_')[0])"}
    remove-appxprovisionedPackage -online -PackageName $photosProvisionedApp.PackageName

    write-output "Provisioning $($appInstaller) to system..."
    Try{
        Add-AppxProvisionedPackage -Online -PackagePath "$($installFile)" -SkipLicense
        Write-Output "successfully provisioned $($appInstaller) to system"
    }catch{
        Write-Output "failed to provision $($appInstaller) to system"
    }
}

Stop-Transcript
EXIT 0