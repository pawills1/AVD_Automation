#determine installation log path;
If(!(test-path -Path "$env:SystemDrive\Program Files\AVDAppInstallLogs")){
    New-Item -ItemType Directory -Path "$env:SystemDrive\Program Files\AVDAppInstallLogs" -Force
}

$logFolder = "$env:SystemDrive\Program Files\AVDAppInstallLogs"
$appNameForLog = "TeamsOutlookAddIn"
$LogPath = "$($logFolder)\$($appNameForLog)_AVD_Script.log"

Start-Transcript -Path $logPath

#//get MSTeams install directory under WindowsApps, create msiexec command line and install Teams Outlook Meeting Add-in for all users;
Write-Output "Performing system-wide install of Teams Meeting Add-in for Office..."
$MSTeams = Get-ChildItem -Path "C:\Program Files\WindowsApps" | ? {($_.Name -like "MSTeams_*") -AND (($_.Name -notlike "*1.0.0.0*"))}

#//get binary version of Teams Meeting Add-In, convert to string & edit to make useable for TARGETDIR parameter;
$teamsAddInBinaryVer = Get-AppLockerFileInformation -Path "$($MSTeams.FullName)\MicrosoftTeamsMeetingAddinInstaller.msi" | select -ExpandProperty Publisher | select BinaryVersion
$teamsAddInStringVer = [string]$teamsAddInBinaryVer
$teamsAddInVer       = $teamsAddInStringVer.split('=')[1].replace('}','')

Start-Process -FilePath "C:\Windows\System32\msiexec.exe" -ArgumentList "/I `"$($MSTeams.FullName)\MicrosoftTeamsMeetingAddinInstaller.msi`" ALLUSERS=1 /qn /norestart TARGETDIR=`"C:\Program Files (x86)\Microsoft\TeamsMeetingAddIn\$($teamsAddInVer)`" /L*V `"$($logFolder)\$($appNameForLog)_AVD_AppInstall.log`"" -Wait

Write-Output "Performing system-wide .dll registration of Microsoft.Teams.AddInLoader.dll"
Start-Process -FilePath "C:\windows\System32\regsvr32.exe" -ArgumentList "/n /s /i:user `"C:\Program Files (x86)\Microsoft\TeamsMeetingAddIn\$($teamsAddInVer)\x64\Microsoft.Teams.AddinLoader.dll`""

Stop-Transcript