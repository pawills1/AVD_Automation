# Function to add entries to hosts file
function Add-Host {
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $Line
    )

    $hostsFile = "$env:windir\System32\drivers\etc\hosts"
    $hosts = Get-Content -Path $hostsFile

    if ($hosts -notcontains $Line) {
        Add-Content -path $hostsFile -value "$Line"
        Write-output "$Line has been added to the hosts file"
    }
    else {
        Write-output "$Line already present in hosts file"
    }
}

#####################################################################################

$workDir = "C:\temp\bootstrap\PrepWinget"
$useZIPfile = $false
$64Bit = [Environment]::Is64BitOperatingSystem

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
If (!(test-path -path "$env:SystemDrive\Program Files\AVDAppInstallLogs")){
    New-Item -ItemType Directory -Path "$env:SystemDrive\Program Files\AVDAppInstallLogs" -Force
}

$logFolder =  "$env:SystemDrive\Program Files\AVDAppInstallLogs"
$appNameForLog = $($workDir).split('\')[3]
$LogPath = "$($logFolder)\$($appNameForLog)_AVD_Script.log"

Start-Transcript -Path $logPath

$InstallerPath = Get-childitem -path . -File

If($useZIPfile -eq $TRUE){
    Write-Output "Extracting contents of $($ZIPFile) to $($workDir)..."
    $ZIPFile = $InstallerPath | Where-Object {$_.Extension -match ".zip"}
    Expand-Archive -literalPath $ZIPFile -DestinationPath $workDir -Force
}

# Check winget installation status
Write-output 'Testing whether winget is installed.'
$wingetStatus = Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -eq 'Microsoft.DesktopAppInstaller'}

if ([Version]$wingetStatus.Version -ge "2023.808.2243.0") {
	Write-output "WinGet is Installed with an acceptable version."
}
else {
	# Download the latest version of winget
    Write-output 'Winget is not installed or not up to date. Downloading latest version.'
    $source = 'https://aka.ms/getwinget'
    Invoke-WebRequest -Uri $source -OutFile "$workDir\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
    Unblock-File -Path "$workDir\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"

	# Install the downloaded msixbundle
	try {
		write-output 'Installing MSIXBundle for the Desktop App Installer.'
		Add-AppxProvisionedPackage -Online -PackagePath "$workDir\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle" -SkipLicense
		write-output 'Installed MSIXBundle for the Desktop App Installer.'
	}
	catch {
		write-output 'Installation failed'
	}
}

# Check vcredist installation status
if ($64Bit -eq $true) {
    $vcStatus = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\X64' -Name 'Installed' -ErrorAction SilentlyContinue).Installed
}
else {
    $vcStatus = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\X86' -Name 'Installed' -ErrorAction SilentlyContinue).Installed
}

Write-output 'Testing whether Microsoft Visual C++ Redistributable is installed.'

if ($vcStatus -eq '1') {
	Write-output "The appropriate runtime is installed."
}
else {
	# Download and install the latest version of  the runtime
    if ($64Bit -eq $true) {
        Write-output 'The visual C runtime is not installed. Downloading latest version.'
        $source = 'https://aka.ms/vs/17/release/vc_redist.x64.exe'
        Invoke-WebRequest -Uri $source -OutFile "$workDir\vc_redist.x64.exe"
        Unblock-File -Path "$workDir\vc_redist.x64.exe"

        try {
		    Write-output 'Installing the runtime.'
            Start-Process -Wait -FilePath "$workDir\vc_redist.x64.exe" -ArgumentList "/q /norestart"
		    Write-output 'Runtime installed!'
	    }
	    catch {
		    Write-output 'Installation failed'
	    }
    }
    else {
        Write-output 'The visual C runtime is not installed. Downloading latest version.'
        $source = 'https://aka.ms/vs/17/release/vc_redist.x86.exe'
        Invoke-WebRequest -Uri $source -OutFile "$workDir\vc_redist.x86.exe"
        Unblock-File -Path "$workDir\vc_redist.x86.exe"

        try {
		    write-output 'Installing the runtime.'
            Start-Process -Wait -FilePath "$workDir\vc_redist.x86.exe" -ArgumentList "/q /norestart"
		    write-output 'Runtime installed!'
	    }
	    catch {
		    write-output 'Installation failed'
	    }
    }
}

# Add winget dev repo certificate to Trusted Root Store
Write-output 'Installing Winget Dev Certificate.'
Import-Certificate -Filepath "$($workdir)\testCA.crt" -CertStoreLocation 'Cert:\LocalMachine\Root'

# Add entries required for the dev repo to hosts file
$hostsToAdd = @("172.174.67.175`twingethost.army.mil","20.159.193.177`twinget-sbx.army.mil","52.227.141.18`twinget-dev.army.mil")

foreach ($item in $hostsToAdd) {
    Add-Host -Line $item
}

# Add the Army private repo. Currently using dev
if (!(Test-Path 'HKLM:SOFTWARE\Policies\Microsoft\Windows\AppInstaller')) {
    New-Item -Path 'HKLM:SOFTWARE\Policies\Microsoft\Windows\' -Name 'Appinstaller'
}

if (!(Test-Path 'HKLM:SOFTWARE\Policies\Microsoft\Windows\AppInstaller\AdditionalSources')) {
    New-Item -Path 'HKLM:SOFTWARE\Policies\Microsoft\Windows\AppInstaller\' -Name 'AdditionalSources'
}

Set-ItemProperty -Path 'HKLM:SOFTWARE\Policies\Microsoft\Windows\AppInstaller\' -Name 'EnableAdditionalSources' -Value 1
Set-ItemProperty -Path 'HKLM:SOFTWARE\Policies\Microsoft\Windows\AppInstaller\AdditionalSources' -Name '1' -Value '{"Arg":"https://winget-dev.army.mil/api/","Data":"","Explicit":false,"Identifier":"winget-cli-restsource-dev","Name":"dev","TrustLevel":["Trusted"],"Type":"Microsoft.Rest"}'

# Configure scheduled task for daily updates of winget apps that are installed in the system context.
$taskXml = @'
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Author>U.S. Army NETCOM</Author>
    <URI>\WingetUpdateSystemApps</URI>
  </RegistrationInfo>
  <Principals>
    <Principal id="Author">
      <UserId>S-1-5-18</UserId>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
    <UseUnifiedSchedulingEngine>true</UseUnifiedSchedulingEngine>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT4H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>PowerShell.exe</Command>
      <Arguments>Invoke-Command -ScriptBlock {
$winget = $null
$systemContext = Resolve-Path "${env:ProgramFiles}\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe"

if ($systemContext.count -gt 1) {
    $systemContext = $systemContext[-1].Path
}

if (Test-Path "$systemContext\winget.exe") {
    $winget = (Get-ChildItem -Path $systemContext | Where-Object -FilterScript {$_.Name -like 'winget.exe'}).FullName
}
else {
    Throw 'Winget is not installed and the machine requires preparation.'
}

Start-Process $winget -Wait -ArgumentList 'upgrade --all --source dev --scope machine --accept-package-agreements --accept-source-agreements --disable-interactivity --silent'}</Arguments>
    </Exec>
  </Actions>
</Task>
'@

Write-Output 'Creating scheduled task for Updating Winget apps in System Context'
Write-Output 'Checking for existing scheduled task for Updating Winget apps in System Context'

$scheduledTask = Get-ScheduledTask -TaskName 'WingetUpdateSystemApps' -ErrorAction SilentlyContinue

if ($null -eq $scheduledTask) {
    write-output 'Scheduled task does not exist, attempting to create new sceduled task'
    try {
        Register-ScheduledTask -Xml $taskXml -TaskName 'WingetUpdateSystemApps' -ErrorAction Stop
        write-output 'Created new WingetUpdateSystemApps scheduled task'
    }
    catch {
        write-output 'Error creating WingetUpdateSystemApps scheduled task'
        write-output $Error[0]
    }
}
else {
    write-output 'Existing WingetUpdateSystemApps scheduled task exists, deleting existing one first'
    try {
        Unregister-ScheduledTask -TaskName 'WingetUpdateSystemApps' -Confirm:$false -ErrorAction Stop
        Register-ScheduledTask -Xml $taskXml -TaskName 'WingetUpdateSystemApps' -ErrorAction Stop
        write-output 'Created a replacement WingetUpdateSystemApps scheduled task'
    }
    catch {
        write-output 'Error removing and creating the WingetUpdateSystemApps scheduled task'
        write-output $Error[0]
    }
}

Write-Output 'Script finished'
Stop-Transcript


