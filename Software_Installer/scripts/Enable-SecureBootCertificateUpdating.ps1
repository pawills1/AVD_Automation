$STXMLContents = @'
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>2026-03-10T10:51:07.9492244</Date>
    <Author>ARMY</Author>
    <Description>Updating Secure Boot DB Certificates prior to expiration in 06/2026.</Description>
    <URI>\Army AVD - Secure Boot DB Certificate Updates</URI>
  </RegistrationInfo>
  <Triggers>
    <BootTrigger>
      <Enabled>true</Enabled>
    </BootTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>S-1-5-18</UserId>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>true</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>false</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT72H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe</Command>
      <Arguments>-ExecutionPolicy Bypass -File "C:\Windows\SysWOW64\armyAVDdsc\Scripts\Run-SecureBootDBCertificateUpdates.ps1"</Arguments>
    </Exec>
  </Actions>
</Task>
'@

#//enable system for Secure Boot DB Certificates if not already enabled;
If ((&WinCSFlags.exe /query --key F33E0C8E002) -like "*Disabled*"){

    Write-Output "System not configured for Secure Boot DB Certificate Updates. Configuring it now."
    &winCSFlags.exe /apply --key F33E0C8E002
    <#
    #//set triggers in the Secure-Boot-Update ST to false, so that this sched task can be run manually;
    Get-ScheduledTask -TaskName "Secure-Boot-Update" | Export-ScheduledTask | Out-File "C:\Windows\Temp\Working.xml" -Force
    [xml]$xmlContent = Get-Content -Path "C:\Windows\Temp\Working.xml" -raw

    $xmlContent.Task.Triggers.BootTrigger.Enabled = "false"
    $xmlContent.Save("C:\Windows\Temp\Working.xml")

    Unregister-ScheduledTask -TaskName "Secure-Boot-Update" -Confirm:$false
    schtasks.exe /Create /XML C:\Windows\Temp\Working.xml /tn "\Microsoft\Windows\PI\Secure-Boot-Update"
    Remove-Item -Path "C:\Windows\Temp\Working.xml" -Force

    #//create new scheduled task to run the secre-boot-update scheduled task manually twice, focrcing system to prepare itself to update certs on next reboot;
    #//This ST runs a script called Run-SecureBootUpdateCertificates.ps1 which will delete this ST upon successful completion;
    Register-ScheduledTask -Xml $STXMLContents -TaskName 'Army AVD - Secure Boot DB Certificate Updates' | Out-Null
    #>
}else{

    Write-Output "System already configured for Secure Boot DB Certificate Updates."
}
