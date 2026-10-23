##########################################################################################
#
# FORCES OS BACK TO USING NON-PREVIEW BUILD OF WINDOWS API;
# https://learn.microsoft.com/en-sg/answers/questions/1791668/powershell-command-outputting-system-comobject-on
# https://github.com/rgl/packer-plugin-windows-update/issues/144
#
##########################################################################################

$wuVer = (New-Object -ComObject Microsoft.Update.AgentInfo).GetInfo('ProductVersionString')

$path = "C:\ProgramData\Microsoft\Windows\OneSettings"
$file = "UusSettings.json"
$renamedFile = "UusSettings.json.BAK"

If(Test-Path -Path $path){
    Write-Output "$($path) found. Checking for existence of $($file)..."

    If(Test-Path -Path "$($path)\$($file)"){
        $uusObj = Get-Content "$($path)\$($file)" | ConvertFrom-Json

        if ($uusObj.settings.EXCLUSIONS -notcontains $wuVer){
            Write-Output "Windows Update API version $($wuVer) not found in $($file). Updating file with current WU API Version..."

            Copy-Item -Path "$($path)\$($file)" -Destination "$($path)\$($renamedFile)" -force
            $uusObj.settings.EXCLUSIONS += "$wuVer"
            " " | Out-File "$($path)\$($file)" -Encoding ascii -NoNewline -Force
            $uusObj | ConvertTo-Json -Compress | Out-File "$($path)\$($file)" -Encoding UTF8 -Append -Force
            
        }Else{
            Write-Output "$($file) already contains WU API version $($wuVer). No need to edit it!"
        }

    }Else{
        Write-Output "Could not locate $($file) in $($path)"
    }

}Else{
    Write-Output "Could not locate $($path) on this system!"
}