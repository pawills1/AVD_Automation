$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

If (gci -Path 'C:\ProgramData\Microsoft\Windows\AppRepository' | ? {$_.name -like 'StateRepository-*'}){
    Write-Output "Removing AppXStateRepository backup folder. Appx StateRepository recreated on host..."
    Remove-Item -Path 'C:\temp\StateReadinessBkup'-Force
}else{
    throw "State Repository NOT re-created on host!"
}