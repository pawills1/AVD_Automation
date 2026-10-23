$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

$regKey = "HKLM:\SOFTWARE\ArmyAVD"
$arrRegVal = @()
$arrRegVal = $arrRegVal + [pscustomobject]@{RegName = "ImageDefinition" ; RegVal = "$($env:ImgDefName)"}
$arrRegVal = $arrRegVal + [pscustomobject]@{RegName = "Sprint" ; RegVal = (get-date -format 'MM-yyyy')}
$arrRegVal = $arrRegVal + [pscustomobject]@{RegName = "ImageVersion" ; RegVal = "$($env:ImgVersion)"}
$arrRegVal = $arrRegVal + [pscustomobject]@{RegName = "TimeStamp" ; RegVal = "{0:MM/dd/yy} {0:HH:mm:ss} UTC" -f (Get-Date).ToUniversalTime()}


If(!(test-path -Path $regKey -ErrorAction SilentlyContinue)){
    New-Item -Path $regKey -Force
}

Write-Output "Writing image build information into $($regKey)..."

ForEach($setting in $arrRegVal){
    Set-ItemProperty -Path $regKey -Name $setting.regName -Value $setting.RegVal
}
