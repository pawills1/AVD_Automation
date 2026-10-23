$NotepadDirectories = @{
    Source      = "C:\temp\bootstrap\Notepad++"
    Destination = "C:\Notepad++"
}
Move-Item -Path $NotepadDirectories.Source -Destination $NotepadDirectories.Destination -Force