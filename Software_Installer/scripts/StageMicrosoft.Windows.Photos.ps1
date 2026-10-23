$FileDirectories = @{
    Source      = "C:\temp\bootstrap\Microsoft.Windows.Photos\Microsoft.Windows.Photos_AVD.zip"
    Destination = "C:\Windows\SysWOW64\ArmyAVDdsc\Applications\Microsoft.Windows.Photos_AVD.zip"
}
Move-Item -Path $FileDirectories.Source -Destination $FileDirectories.Destination -Force