New-Item -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' -Force
New-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' -name 'DeferQualityUpdates' -Value 1 -PropertyType dword -Force
New-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' -name 'DeferQualityUpdatesPeriodInDays' -Value 30 -PropertyType dword -Force
New-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' -name 'PauseQualityUpdatesStartTime' -Value "2025-07-10" -PropertyType string -Force