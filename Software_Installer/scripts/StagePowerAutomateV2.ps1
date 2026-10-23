$PowerAutomateDirectories = @{
    Source      = "C:\temp\bootstrap\PowerAutomate\PowerAutomate_AVD.zip"
    Destination = "C:\Windows\SysWOW64\ArmyAVDdsc\Applications\PowerAutomate"
}
Move-Item -Path $PowerAutomateDirectories.Source -Destination $PowerAutomateDirectories.Destination  -Force