remove-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' -name 'DeferQualityUpdates' -Force
remove-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' -name 'DeferQualityUpdatesPeriodInDays' -Force
remove-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' -name 'PauseQualityUpdatesStartTime' -Force