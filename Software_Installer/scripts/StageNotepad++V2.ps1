$NotepadDirectories = @{
    Source      = "C:\temp\bootstrap\Notepad++\Notepad++_AVD.zip"
    Destination = "C:\Windows\SysWOW64\ArmyAVDdsc\Applications\Notepad++"
}
Move-Item -Path $NotepadDirectories.Source -Destination $NotepadDirectories.Destination -Force