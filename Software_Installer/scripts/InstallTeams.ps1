$install = $true
$workDir = "C:\temp\bootstrap\Teams"
$useZIPfile = $false

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
$LogPath = "$($logFolder)\$($appNameForLog)_AVD_Install.log"

Start-Transcript -Path $logPath

$InstallerPath = Get-childitem -path . -File

If($useZIPfile -eq $TRUE){
    Write-Output "Extracting contents of $($ZIPFile) to $($workDir)..."
    $ZIPFile = $InstallerPath | Where-Object {$_.Extension -match ".zip"}
    Expand-Archive -literalPath $ZIPFile -DestinationPath $workDir -Force
}Else{
    Invoke-WebRequest -Uri "https://go.microsoft.com/fwlink/?linkid=2196106" -OutFile "$($workdir)\MSTeams-x64.msix"
    Invoke-WebRequest -Uri "https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409" -OutFile "$($workdir)\teamsbootstrapper.exe"
}

#Detect Installer
$InstallerPath = Get-childitem -path . -File

#//dynamic detect .MSIX file;
$installFile = $InstallerPath | Where-Object {$_.Extension -match ".exe"}
$MSIXFile = $InstallerPath | Where-Object {$_.Extension -match ".msix"}

#//ensure disable auto-update in place before Teams install...
Write-Output "creaating new Teams registry key and setting to disable auto updates..."
If (!(Get-ItemProperty -path 'HKLM:\SOFTWARE\Microsoft\Teams' -ErrorAction SilentlyContinue)){
    New-Item -Path 'HKLM:\SOFTWARE\Microsoft' -Name 'Teams' -Force -ErrorAction SilentlyContinue | Out-Null
}
New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Teams' -Name "disableAutoUpdate" -Value 1 -PropertyType DWORD -Force -ErrorAction SilentlyContinue | Out-Null

#Determine if EXE or MSI - use specific installer switches for each.
if ($installFile -match "exe") {
    # Firefox specific arguments
    Write-Output "EXE found, beginning installation with provided switches."
    Start-Process -FilePath $installFile.FullName -ArgumentList "-p -o `"$($MSIXFile.FullName)`"" -Wait
}
else {
    #Assume MSI - standard silent arguments
    Write-Output "Not an EXE, assumed MSI, beginning install with silent switches"
    Start-Process -FilePath $installFile.FullName -ArgumentList "/qn" -Wait
} 

<#
#//get MSTeams install directory under WindowsApps, create msiexec command line and install Teams Outlook Meeting Add-in for all users;
Write-Output "Performing system-wide install of Teams Meeting Add-in for Office..."
$MSTeams = Get-ChildItem -Path "C:\Program Files\WindowsApps" | ? {$_.Name -like "MSTeams_*"}

#//get binary version of Teams Meeting Add-In, convert to string & edit to make useable for TARGETDIR parameter;
$teamsAddInBinaryVer = Get-AppLockerFileInformation -Path "$($MSTeams.FullName)\MicrosoftTeamsMeetingAddinInstaller.msi" | select -ExpandProperty Publisher | select BinaryVersion
$teamsAddInStringVer = [string]$teamsAddInBinaryVer
$teamsAddInVer       = $teamsAddInStringVer.split('=')[1].replace('}','')

Start-Process -FilePath "C:\Windows\System32\msiexec.exe" -ArgumentList "/I `"$($MSTeams.FullName)\MicrosoftTeamsMeetingAddinInstaller.msi`" ALLUSERS=1 /qn /norestart TARGETDIR=`"C:\Program Files (x86)\Microsoft\TeamsMeetingAddIn\$($teamsAddInVer)`" /L*V `"$($logFolder)\TeamsMeetingAddInAppInstaller.log`"" -Wait
#>
Stop-Transcript