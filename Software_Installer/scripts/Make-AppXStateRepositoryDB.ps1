$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

 New-Item -Path 'C:\temp\StateReadinessBkup' -Force
 
Stop-Service AppXSvc -Force
Stop-Service StateRepository -Force
Stop-Service AppReadiness -Force

cd 'C:\ProgramData\Microsoft\Windows\AppRepository'

Write-Output "Moving state repository files from C:\ProgramData\Microsoft\Windows\AppRepository into C:\temp\StateReadinessBkup..."
gci -Path 'C:\ProgramData\Microsoft\Windows\AppRepository' | ? {$_.name -like 'StateRepository-*'} | % {Move-Item -Path $_.name -Destination 'C:\temp\StateReadinessBkup' -Force}

If(gci -Path 'C:\temp\StateReadinessBkup' | ? {$_.name -like 'StateRepository-*'}){
    Write-Output "State Repository files have been successfully moved into C:\temp\StateReadinessBkup"
}else{
    throw "State Repository files WERE NOT successfully moved!"
}
