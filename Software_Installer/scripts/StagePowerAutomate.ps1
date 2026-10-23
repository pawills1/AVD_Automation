$PowerAutomateDirectories = @{
    Source      = "C:\temp\bootstrap\PowerAutomate"
    Destination = "C:\PowerAutomate"
}
Move-Item -Path $PowerAutomateDirectories.Source -Destination $PowerAutomateDirectories.Destination  -Force