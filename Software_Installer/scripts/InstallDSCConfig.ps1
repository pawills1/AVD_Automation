$install = $true
$workDir = "C:\temp\bootstrap\DSCConfig"
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

#//create C:\Program Files\ArmyAVD directory if not there;
Write-Output "creating c:\Windows\SysWOW64\ArmyAVDdsc"
If (!(Test-Path -Path "c:\Windows\SysWOW64\ArmyAVDdsc")){
    mkdir "c:\Windows\SysWOW64\ArmyAVDdsc" -ErrorAction SilentlyContinue
}
Write-Output "creating c:\Windows\SysWOW64\ArmyAVDdsc\Applications"
If (!(Test-Path -Path "c:\Windows\SysWOW64\ArmyAVDdsc\Applications")){
    mkdir "c:\Windows\SysWOW64\ArmyAVDdsc\Applications" -ErrorAction SilentlyContinue
}
Write-Output "creating c:\Windows\SysWOW64\ArmyAVDdsc\Applications\Notepad++"
If (!(Test-Path -Path "c:\Windows\SysWOW64\ArmyAVDdsc\Applications\Notepad++")){
    mkdir "c:\Windows\SysWOW64\ArmyAVDdsc\Applications\Notepad++" -ErrorAction SilentlyContinue
}
Write-Output "creating c:\Windows\SysWOW64\ArmyAVDdsc\Applications\PowerAutomate"
If (!(Test-Path -Path "c:\Windows\SysWOW64\ArmyAVDdsc\Applications\PowerAutomate")){
    mkdir "c:\Windows\SysWOW64\ArmyAVDdsc\Applications\PowerAutomate" -ErrorAction SilentlyContinue
}
Write-Output "creating c:\Windows\SysWOW64\ArmyAVDdsc\Scripts"
If (!(Test-Path -Path "c:\Windows\SysWOW64\ArmyAVDdsc\Scripts")){
    mkdir "c:\Windows\SysWOW64\ArmyAVDdsc\Scripts" -ErrorAction SilentlyContinue
}
Write-Output "creating c:\Windows\SysWOW64\ArmyAVDdsc\Policies"
If (!(Test-Path -Path "c:\Windows\SysWOW64\ArmyAVDdsc\Policies")){
    mkdir "c:\Windows\SysWOW64\ArmyAVDdsc\Policies" -ErrorAction SilentlyContinue
}


#//copy items into C:\Program Files\ArmyAVD, theme & wallpaper directores;
write-output "Copying files into c:\Windows\SysWOW64\ArmyAVDdsc"
copy-item "$($workdir)\localhost.mof" -destination "c:\Windows\SysWOW64\ArmyAVDdsc" -Force
copy-item "$($workdir)\localhost.meta.mof" -destination "c:\Windows\SysWOW64\ArmyAVDdsc" -Force 
copy-item "$($workdir)\DSC-read.ps1" -destination "c:\Windows\SysWOW64\ArmyAVDdsc" -Force
copy-item "$($workdir)\DSC.xml" -destination "c:\Windows\SysWOW64\ArmyAVDdsc" -Force
copy-item "$($workdir)\CDrivePermissions-Remediate.ps1" -destination "c:\Windows\SysWOW64\ArmyAVDdsc\Scripts" -Force
copy-item "$($workdir)\PDrivePermissions-Remediate.ps1" -destination "c:\Windows\SysWOW64\ArmyAVDdsc\Scripts" -Force
copy-item "$($workdir)\DDrivePermissions-Remediate.ps1" -destination "c:\Windows\SysWOW64\ArmyAVDdsc\Scripts" -Force
copy-item "$($workdir)\UnapprovedStoreApps-DetectAndRemediate.ps1" -destination "c:\Windows\SysWOW64\ArmyAVDdsc\Scripts" -Force
copy-item "$($workdir)\DSCInitialize.ps1" -destination "c:\Windows\SysWOW64\ArmyAVDdsc\Scripts" -Force
copy-item "$($workdir)\NessusFolderPermissions-Remediate.ps1" -destination "c:\Windows\SysWOW64\ArmyAVDdsc\Scripts" -Force
copy-item "$($workdir)\InstallNotepad++.ps1" -destination "c:\Windows\SysWOW64\ArmyAVDdsc\Applications\Notepad++" -Force
copy-item "$($workdir)\InstallPowerAutomate.ps1" -destination "C:\Windows\SysWOW64\ArmyAVDdsc\Applications\PowerAutomate" -Force
<#
copy-item "$($workdir)\CIActivatePolicies.ps1" -destination "c:\Windows\SysWOW64\ArmyAVDdsc\Scripts" -Force
copy-item "$($workdir)\Run-SecureBootDBCertificateUpdates.ps1" -destination "c:\Windows\SysWOW64\ArmyAVDdsc\Scripts" -Force
copy-item "$($workdir)\{5B4CDE98-6AA7-4043-B3C2-98112B6BF091}.cip" -destination "c:\Windows\SysWOW64\ArmyAVDdsc\Policies" -Force
copy-item "$($workdir)\{20DA04D0-D83F-4ACB-93AF-FCF9E8F62B9D}.cip" -destination "c:\Windows\SysWOW64\ArmyAVDdsc\Policies" -Force
copy-item "$($workdir)\{5392B3ED-10E2-4047-A8B1-67F0031E1B1A}.cip" -destination "c:\Windows\SysWOW64\ArmyAVDdsc\Policies" -Force
copy-item "$($workdir)\{0AE05A8A-5635-4D61-8AB5-E3E26CBF6AFE}.cip" -destination "c:\Windows\SysWOW64\ArmyAVDdsc\Policies" -Force
#>

#//Create Army AVD Event View Application Source, if not found;
Write-Output "Checking if Army AVD Log Source exists"
if (!([System.Diagnostics.EventLog]::SourceExists("Army AVD")))
{
    Write-Output "Army AVD Log Source does not exist, attempting to create and creating event (ID 5000) to record creation..."
    New-EventLog -LogName Application -source "Army AVD"
    Write-EventLog -LogName Application -Source "Army AVD" -EventID 5000 -EntryType Information -Message "EventLot Army AVD Initialized" -Category 0
}
else
{
    Write-Output "Army AVD Log Source found"
}

#//Run enable-PSRemoting cmdlet to create FW rule for WinRM HTTP management (por5985);
Write-Output "WinRM not configured to use port 5985, confinguring FW rule to allow for WinRM HTTP management..."
Enable-PSRemoting -Force

#//check for state of DSC LCM ConfigurationMode property. If not set to ApplyAndMontor, set it to this value;
If (!((Get-DscLocalConfigurationManager).ConfigurationMode -eq "ApplyAndAutoCorrect")){
    Write-output "Changing DSC LCM ConfigurationMode property to ApplyAndAutoCorrect..."
    Set-DscLocalConfigurationManager -Path "c:\Windows\SysWOW64\ArmyAVDdsc" -Force
}

#//apply AVD DSC configuration;
Try{
    Write-Output "Applying Army AVD MOF file..."
    Start-DSCConfiguration -path "c:\Windows\SysWOW64\ArmyAVDdsc" -wait -force
    Write-EventLog -LogName Application -Source "Army AVD" -EventID 5001 -EntryType Information -Message "Successfully applied AVD - DSC Config" -Category 0
    Write-Output "Successfully applied AVD MOF file..."
}Catch{
    Write-Output "Unable to apply AVD MOF file..."
    Write-EventLog -LogName Application -Source "Army AVD" -EventID 5559 -EntryType Error -Message "Could not apply AVD - DSC Config" -Category 0
}

#//create AVD DSC scheduled task;
write-output "creating AVD DSC Read scheduled task..."
schtasks.exe /Create /XML c:\Windows\SysWOW64\ArmyAVDdsc\DSC.xml /tn "AVD DSC Read"

#//change path back to location before $workdir set;
If($workDir){
    Set-Location -Path $origLocation
}

Stop-Transcript

EXIT 0
