#Task Manager Test AVD-00
$taskManager = Start-Process taskmgr
Write-Host -ForegroundColor Cyan "Opening Task Manager"

#7-Zip Test AVD-02
Write-Host -ForegroundColor Cyan "Running 7-Zip Test...Creating UATtest.txt if it doesn't exist and sending to UATtest.Zip..."
$oneDriveDesktopPath = Set-location "$($env:USERPROFILE)\OneDrive - US Army\Desktop"
$oneDriveDesktopTxtPath = "$($env:USERPROFILE)\OneDrive - US Army\Desktop\UATtest.txt"
$oneDriveDesktopZipPath = "$($env:USERPROFILE)\OneDrive - US Army\Desktop\UATtest.zip"

if (Test-Path $oneDriveDesktopTxtPath)
{
Write-Host -ForegroundColor Cyan "UATtest.txt exists"
}
else
{
Write-Host -ForegroundColor Cyan "UATtest.txt doesn't exist creating UATtest.txt"
New-Item "UATtest.txt"
}

if (Test-Path $oneDriveDesktopZipPath)
{
Write-Host -ForegroundColor Cyan "UATtest.zip exists removing UATtest.zip"
Remove-Item "UATtest.zip"
}

$7zipPath = "$env:ProgramFiles\7-Zip\7z.exe"
if (-not (Test-Path -Path $7zipPath -PathType Leaf)) {
    throw "7 zip file '$7zipPath' not found"}

Set-Alias Start-SevenZip $7zipPath
$Source = "$($env:USERPROFILE)\OneDrive - US Army\Desktop\UATtest.txt"
$Target = "$($env:USERPROFILE)\OneDrive - US Army\Desktop\UATtest.zip"
Start-SevenZip -mx=9 a $Target $Source

If(Test-Path $Target -PathType Leaf)
{
Write-Host -ForegroundColor Green "UATtest.zip was created and exists on desktop 7-Zip test PASSED"
}
else
{
Write-Host -ForegroundColor Red "UATtest.zip does not exist on desktop 7-Zip test FAILED"
}

#Adobe Acrobat AVD-03 - AVD-04
$adobeAcrobat = Start-Process "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
Write-Host -ForegroundColor Cyan "Opening Adobe Acrobat"

Write-Host -ForegroundColor Yellow "Pause for manual check of Adobe Acrobat tests AVD-02 - AVD-03, press enter after tests are complete"
Read-Host #press button
Write-Host -ForegroundColor Cyan "Continuing Tests"

#AWSWickr AVD-05
$awsWickr = Start-Process "C:\Program Files\Amazon Web Services, Wickr\AWS WickrGov\AWSWickrGov.exe"
Write-Host -ForegroundColor Cyan "Opening AWSWickrGov"
Write-Host -ForegroundColor Yellow "Pause for manual check of AWSWickr tests AVD-05, press enter after tests are complete"
Read-Host #press button

Write-Host -ForegroundColor Yellow "Pause for manual check of Classification Banner test AVD-06, press enter after tests are complete"
Read-Host #press button

#DAX Studio AVD-07
$daxStudio = Start-Process "c:\Program Files\DAX Studio\DaxStudio.exe"
Write-Host -ForegroundColor Cyan "Opening DAX Studio"
Write-Host -ForegroundColor Yellow "Pause for manual check of DAX Studio tests AVD-07, press enter after tests are complete"
Read-Host #press button


#DEPTH AVD-08
Write-Host -ForegroundColor Cyan "DEPTH Test Spotcheck Keys"
$depth = Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" | Where-Object {$_.Name -like "*Depth*"}
$depth.Name

Write-Host -ForegroundColor Yellow "Pause for manual check of DEPTH Keys test AVD-08, press enter after tests are complete"
Read-Host #press button

#DTP Lite AVD-09
$dtpLite = Start-Process "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\DTP Lite\DTP Lite"
Write-Host -ForegroundColor Cyan "Opening DTPLite"
Write-Host -ForegroundColor Yellow "Pause for manual check of DTP Lite tests AVD-09, press enter after tests are complete"
Read-Host #press button

#Foundry ODBC Driver AVD-10
$foundryODBCPath = "C:\Program Files\Palantir\Foundry ODBC Driver"

Write-Host -ForegroundColor Cyan "Checking ODBC Driver install path (UAT test AVD-10)"

If(Test-Path -Path $foundryODBCPath) 
{
Write-Host -ForegroundColor Green "Foundry ODBC Driver is Installed, AVD-10 test passed"
}
else
{
Write-Host -ForegroundColor Red "Foundry ODBC Driver is NOT Installed, test FAILED"
}

Write-Host -ForegroundColor Yellow "Pause for manual check of Foundry ODBC Driver tests AVD-10, see results for this test above, press enter after tests are complete"
Read-Host #press button

#Google Chrome Tests AVD-11 - AVD-16
Write-Host -ForegroundColor Cyan "Google Chrome Tests AVD-11 - AVD-16"
Write-Host -ForegroundColor Cyan "Opening Google Chrome"
$chromePath = Start-Process "chrome" "https://youtube.com"
Write-Host -ForegroundColor Cyan "Chrome Policy Check AVD-16"
$chromePolicyRegPath = "HKLM:\SOFTWARE\Policies\Google\Chrome"

$expected = @{
    "HistorySearchSettings"              = 2
    "CreateThemesSettings"               = 2
    "DevToolsGenAISettings"             = 2
    "AutoFillPredictionSettings"        = 2
    "GenAILocalFoundationalModelSettings" = 1
    "HelpMeWriteSettings"               = 2
    "TabCompareSettings"                = 2
}

if (-not (Test-Path $chromePolicyRegPath)) {
    Write-Host "Chrome policy key not found: $chromePolicyRegPath" -ForegroundColor Red
    return
}

$actual = Get-ItemProperty -Path $chromePolicyRegPath

foreach ($key in $expected.Keys) {
    $expectedValue = $expected[$key]
    $actualValue   = $actual.$key

    if ($null -eq $actualValue) {
        Write-Host "$key : FAIL (expected $expectedValue)" -ForegroundColor Yellow
    }
    elseif ($actualValue -eq $expectedValue) {
        Write-Host "$key : PASS ($actualValue)" -ForegroundColor Green
    }
    else {
        Write-Host "$key : FAIL (current=$actualValue expected=$expectedValue)" -ForegroundColor Red
    }
}

Write-Host -ForegroundColor Yellow "Pause for manual check of Google Chrome tests AVD-11 - AVD-16, press enter after tests are complete"


Read-Host #press button

#Office Tests AVD-18 - AVD-19
Write-Host -ForegroundColor Cyan "Opening Office Apps"
$wordPath = Start-Process "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
$outlookPath = Start-Process "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
$powerPointPath = Start-Process "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
$excelPath = Start-Process "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
$visioPath = Start-Process "C:\Program Files\Microsoft Office\root\Office16\VISIO.EXE"
$projectPath = Start-Process "C:\Program Files\Microsoft Office\root\Office16\WINPROJ.EXE"
$oneNotePath = Start-Process "C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE"

Write-Host -ForegroundColor Yellow "Pause for manual check of Office tests AVD-18 and AVD-19, press enter after tests are complete"
Read-Host #press button

#Teams Test AVD-20 - AVD-24
Write-Host -ForegroundColor Cyan "Opening Teams"
$teams = Start-Process "ms-teams.exe"
Write-Host -ForegroundColor Yellow "Pause for manual check of Teams tests AVD-20 - AVD-24, press enter after tests are complete"
Read-Host #press button

#FireFox Test AVD-25 - AVD-26
Write-Host -ForegroundColor Cyan "Opening Mozilla FireFox"
$fireFoxPathYT = Start-Process firefox "https://youtube.com" 
$fireFoxPathStigs = Start-Process firefox "about:policies"
Write-Host -ForegroundColor Yellow "Pause for manual check of FireFox tests AVD-25 - AVD-26, press enter after tests are complete"
Read-Host #press button

#Edge Test AVD-28 - AVD-37
Write-Host -ForegroundColor Cyan "Opening Microsoft Edge"
$edge = Start-Process "https://youtube.com"
Write-Host -ForegroundColor Yellow "Pause for manual check of Edge tests AVD-28 - AVD-37, press enter after tests are complete"
Read-Host #press button

#Notepad++ Test AVD-38
Write-Host -ForegroundColor Cyan "Opening Notepad++"
$notepadPlus = Start-Process "C:\Program Files\Notepad++\notepad++.exe"
Write-Host -ForegroundColor Yellow "Pause for manual check of Notepad++ AVD-38, press enter after tests are complete"
Read-Host #press button

#OCE Tool AVD-36 - AVD-37
Write-Host -ForegroundColor Cyan "Opening OCE Tool"
$oceTool = Start-Process "C:\ocetool\oce.exe"
Write-Host -ForegroundColor Yellow "Pause for manual check of OCE Tool AVD-36 - AVD-37, press enter after tests are complete"
Read-Host #press button

#Outlook Tests AVD-41 - AVD-42
Write-Host -ForegroundColor Cyan "Opening Outlook"
$outlookPath = Start-Process "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
Write-Host -ForegroundColor Yellow "Pause for manual check of Outlook AVD-41 - AVD-42, press enter after tests are complete"
Read-Host #press button

#Power Automate Tests AVD-40
Write-Host -ForegroundColor Cyan "Opening Power Automate"
$powerAutomate = Start-Process "C:\Program Files (x86)\Power Automate Desktop\dotnet\PAD.Console.Host.exe"
Write-Host -ForegroundColor Yellow "Pause for manual check of Power Automate AVD-40, press enter after tests are complete"
Read-Host #press button

#Power BI Desktop Tests AVD-44
Write-Host -ForegroundColor Cyan "Opening Power BI Desktop"
$powerAutomate = Start-Process "C:\Program Files\Microsoft Power BI Desktop\bin\PBIDesktop.exe"
Write-Host -ForegroundColor Yellow "Pause for manual check of Power BI Desktop AVD-44, press enter after tests are complete"
Read-Host #press button

#Power BI Report Builder Tests AVD-45
Write-Host -ForegroundColor Cyan "Opening Power BI Report Builder"
$powerAutomate = Start-Process "C:\Program Files\Power BI Report Builder\PowerBIReportBuilder.exe"
Write-Host -ForegroundColor Yellow "Pause for manual check of Power BI Report Builder AVD-45, press enter after tests are complete"
Read-Host #press button

#Powershell 7 and Microsoft.Graph.Modules Test AVD-46 - AVD-47
Write-Host -ForegroundColor Cyan "Powershell 7 test"
Write-Host -ForegroundColor Yellow "Pause for manual check of Powershell 7, if Powershell 7 is open AVD-46 Passed, press enter to continue"
Read-Host #press button

Write-Host -ForegroundColor Cyan "Microsoft.Graph.Modules Test"
Write-Host -ForegroundColor Cyan "Retrieving Microsoft.Graph Modules"

#Retrieves Graph Modules
$moduleVersion = '2.25.0'
$arrModNames = @("Microsoft.Graph.Authentication","Microsoft.Graph.Groups","Microsoft.Graph.Users","Microsoft.Graph.DeviceManagement","Microsoft.Graph.DeviceManagement.Administration","Microsoft.Graph.Identity.DirectoryManagement","Microsoft.Graph.Security","Microsoft.Graph.Beta.DeviceManagement")

ForEach($modName in $arrModNames){
    
    If(!(Get-Module -Name $modName)){
	    Import-Module -name $modName -RequiredVersion $moduleVersion -Force
    }

    Get-Module -Name $modName | select Name, version | FL

}

#Pass/Fail check based on modules retrieved 
$allModulesLoaded = $true

foreach ($modName in $arrModNames) {
    if (-not (Get-Module -Name $modName)) {
        $allModulesLoaded = $false
        break
    }
}

if ($allModulesLoaded) {
    Write-Host "All required Graph modules are loaded. Test AVD-47 Passed" -ForegroundColor Green
} else {
    Write-Host "One or more required Graph modules were not loaded. Test AVD-47 Passed" -ForegroundColor Red
}

Write-Host -ForegroundColor Cyan "Opening portal.azure.us to activate PAG-AVD-Intune-Read-Only PIM group "
Write-Host -ForegroundColor Yellow "This test requires the PAG-AVD-Intune-Read-Only PIM to be activated, please activate PIM group"
$intunePIM = Start-Process "https://portal.azure.us"
Write-Host -ForegroundColor Yellow "Pause to ensure PAG-AVD-Intune-Read-Only PIM group is activated, press enter to continue"
Read-Host #press button

$tenantID = 'fae6d70f-954b-4811-92b6-0530d6f84c43'
$env = 'USGovDoD'
$uri = 'https://graph.microsoft.us'
$graphVersion = "beta"

If(!(Get-MgContext)){
    Connect-MgGraph -Environment $env -TenantId $tenantID
}

$configProfID = 'bfd22ead-81ab-49a2-bed9-c8eaa57960a8'
$reqURI = "$uri/$graphVersion/deviceManagement/configurationPolicies('$configProfID')/settings"
$response = Invoke-MgRestMethod -Uri $reqURI -Method GET | ConvertTo-Json -Depth 50 | ConvertFrom-Json
$response.value.settingInstance

#PrinterLogic Test for Edge and Chrome AVD-48
write-host -ForegroundColor Cyan "Opening Printer Logic on Edge and Chrome try to Add printer on both"
$edgePL = Start-Process "https://print.army.mil/"
$chromePL = Start-Process "chrome" "https://print.army.mil"
Write-Host -ForegroundColor Yellow "Pause for manual check of PrinterLogic AVD-48, press enter after tests are complete"
Read-Host #press button

#Purebred Test AVD-49
write-host -ForegroundColor Yellow "Manually Open Purebred App, this script cannot get to the file location of the EXE"
Write-Host -ForegroundColor Yellow "Pause for manual check of PureBred Agent AVD-50, press enter after tests are complete"
Read-Host #press button

#SAPGUI Test AVD-50
write-host -ForegroundColor Cyan "Opening SAPGUI"
$sapGUI = Start-Process "C:\Program Files\SAP\FrontEnd\SAPGUI\SAPlogon.exe"
Write-Host -ForegroundColor Yellow "Pause for manual check of SAPGUI AVD-45, press enter after tests are complete"
Read-Host #press button

#SelfHelp Tests AVD-51 - AVD-53
write-host -ForegroundColor Cyan "Opening SelfHelp"
$selfHelp = Start-Process "C:\Program Files\SelfHelp\AUDS_SelfHelp.exe"
Write-Host -ForegroundColor Yellow "Pause for manual check of SelfHelp tests AVD-51 - AVD-53, press enter after tests are complete"
Read-Host #press button

#Snipping Tool Tests AVD-54 - AVD-55
write-host -ForegroundColor Cyan "Opening Snipping Tool"
$snippingTool = Start-Process snippingtool
Write-Host -ForegroundColor Yellow "Pause for manual check of Snipping Tool tests AVD-54 - AVD-55, press enter after tests are complete"
Read-Host #press button

#StigViewer Tests AVD-56
write-host -ForegroundColor Cyan "Opening Stig Viewer"
$stigViewer = Start-Process "C:\Program Files (x86)\STIG Viewer 3\STIG Viewer 3.exe"
Write-Host -ForegroundColor Yellow "Pause for manual check of Stig Viewer tests AVD-56, press enter after tests are complete"
Read-Host #press button

#YubiKey Tests AVD-57
write-host -ForegroundColor Cyan "Opening Yubikey Manager"
$yubico = Start-Process "C:\Program Files\Yubico\YubiKey Manager\ykman-gui.exe"
Write-Host -ForegroundColor Yellow "Pause for manual check of Yubikey Manager tests AVD-57, press enter after tests are complete"
Read-Host #press button

#AESMP or IDCO Site Tests AVD-58
write-host -ForegroundColor Cyan "Opening Edge and IDCO Site"
$edgeIDCO = Start-Process "https://idco.dmdc.osd.mil/idco"
$edgeAESMP = Start-Process "https://www.aesmp.army.mil"
Write-Host -ForegroundColor Yellow "Pause for manual check of AESMP or IDCO tests AVD-58, attempt to login with smart card, press enter after tests are complete"
Read-Host #press button

#Wallpaper Tests AVD-54
write-host -ForegroundColor Cyan "Wallpaper Test"
Write-Host -ForegroundColor Yellow "Pause for manual check of Wallpaper tests AVD-59, press enter after tests are complete"
Read-Host #press button

#Drive Encryption Tests AVD-60
write-host -ForegroundColor Cyan "Opening File Explorer (This PC)"
$explorer = Start-Process explorer shell:mycomputerfolder
Write-Host -ForegroundColor Yellow "Pause for manual check of Bitlocker tests AVD-60, press enter after tests are complete"
Read-Host #press button

#Drive Permissions Tests AVD-61 - AVD-62
Write-Host -ForegroundColor Cyan "Drive Permissions Test"
Write-Host -ForegroundColor Cyan "Attempting to create txt file on C: and D: Drive"

#C Drive Permission Check
try {
$cDrivePerm = New-Item -Path c:\ -Name "testtxt.txt" -ItemType "file" -ErrorAction Stop
Write-Host -ForegroundColor Red "C: Drive permission test failed"
}
catch{
Write-Host -ForegroundColor Green "Writing to C: Drive is Denied, AVD-61 Test Passed"
}

#D Drive Permission Check
try {
$pDrivePerm = New-Item -Path d:\ -Name "testtxt.txt" -ItemType "file" -ErrorAction Stop
Write-Host -ForegroundColor Red "D: Drive test failed"
}
catch{
Write-Host -ForegroundColor Green "Writing to D: Drive is Denied, AVD-62 Test Passed"
}

Write-Host -ForegroundColor Yellow "Pause for manual check of Drive Permissions tests AVD-61 - AVD-62, press enter after tests are complete"
Read-Host #press button

#MS Store Tests AVD-63
write-host -ForegroundColor Cyan "Opening MS Store"
$msStoreHelp = Start ms-windows-store:
Write-Host -ForegroundColor Yellow "Pause for manual check of MS Store tests AVD-63, press enter after tests are complete"
Read-Host #press button

#Windows Apps Tests AVD-64
Write-Host -ForegroundColor Cyan "Opening Windows Apps"
$explorer = Start-Process explorer.exe
$calculator = Start-Process calc.exe
$notepad = Start-Process notepad.exe
$snippingTool = Start-Process snippingtool.exe
$photosApp = Start-Process ms-photos:
$paint = Start-Process mspaint

Write-Host -ForegroundColor Yellow "Pause for manual check of Windows Apps tests AVD-59, press enter after tests are complete"
Read-Host #press button

#Install .exe and .msi Tests AVD-65
Write-Host -ForegroundColor yellow "Attempt to install a .exe and .msi installer (ask Will if you need one)"
Write-Host -ForegroundColor Yellow "Pause for manual check of Installer tests AVD-65, press enter after tests are complete"
Read-Host #press button

#Install Store App Install Tests AVD-66
Write-Host -ForegroundColor Yellow "Within the previous test, if you were unable to install the .exe and received the error App control for Business blocked this app then this test AVD-65 passes as well, press enter to continue"
Read-Host #press button

#Language Packs Tests AVD-67
Write-Host -ForegroundColor Cyan "Checking Language Packs AVD-67"
Get-InstalledLanguage
Write-Host -ForegroundColor Yellow "Pause for manual check of Language Packs tests AVD-67, press enter after tests are complete"
Read-Host #press button

#Putty Tests AVD-68
Write-Host -ForegroundColor Cyan "Opening Putty Site"
Start-Process "https://putty.org/index.html"
Write-Host -ForegroundColor Yellow "Pause for manual check, if you were unable to install a .msi for test AVD-68 and unable to download from the putty site then this test passes, press enter after tests are complete"
Read-Host #press button

#OneDrive Tests AVD-69 - AVD-70
write-host -ForegroundColor Cyan "Opening OneDrive"
Start-Process "C:\Program Files\Microsoft OneDrive\OneDrive.exe"
Write-Host -ForegroundColor Yellow "Pause for manual check of OneDrive tests AVD-69 - AVD-70, press enter after tests are complete"
Read-Host #press button

#OS Build Tests AVD-71
write-host -ForegroundColor Cyan "Opening OS Build Site"
write-host -ForegroundColor Cyan "Gathering OS Build of Session Host"
Start-Process "https://learn.microsoft.com/en-us/windows/release-health/windows11-release-information"
systeminfo | findstr /B /C:"OS Name" /B /C:"OS Version"
Write-Host -ForegroundColor Yellow "Pause for manual check of OS Build test AVD-71, press enter after tests are complete"
Read-Host #press button

#Guest Account Tests AVD-72
Write-Host -ForegroundColor Cyan "Checking Guest Account Name"
$guestAccount = get-localuser

If($guestAccount.name -contains "xGuest")
{
Write-Host -ForegroundColor Green "xGuest local account exists Test Passed (AVD-72)"
}
else
{
Write-Host -ForegroundColor Red "xGuest local account does not exist Test Failed"
}

If($guestAccount.name -contains "Guest")
{
Write-Host -ForegroundColor Red "Guest local account Exists Test Failed"
}
else
{
Write-Host -ForegroundColor Green "Guest local account does not exist Test Passed (AVD-72)"
}

Write-Host -ForegroundColor Yellow "Pause for manual check of Guest Account tests AVD-72, press enter after tests are complete"
Read-Host #press button

#Windows Update Tests AVD-73 - AVD-74
Write-Host -ForegroundColor Cyan "Opening Windows Update Settings"
Start-Process explorer ms-settings:windowsupdate
Write-Host -ForegroundColor Yellow "Pause for manual check of Windows Update Settings AVD-73,press enter after tests are complete"
Read-Host #press button
Write-Host -ForegroundColor Cyan "Click Advanced Options within Windows Update"
Start-Process explorer ms-settings:windowsupdate
Write-Host -ForegroundColor Yellow "Pause for manual check of Windows Update Advanced Options AVD-74, press enter after tests are complete"
Read-Host #press button

#Website DNS Tests AVD-70 - AVD-73
Write-Host -ForegroundColor Cyan "DNS Resolution Tests AVD-75 - AVD-78"
$sitesArray = @("idco.dmdc.osd.mil","gears.army.mil","fmsweb.fms.army.mil","prdweb.afmiscloud.army.mil","atrrs.army.mil","stgearsfileprod01.blob.core.usgovcloudapi.net")

foreach($site in $sitesArray)
{
$siteResolution = Resolve-DnsName $site -ErrorAction Ignore
if($siteResolution)
{
Write-Host -ForegroundColor Green "Resolution for $site Succeeded"
}
else
{
Write-Host -ForegroundColor Red "Resolution for $site failed"
}
}

#DNS Alias Test for FMSWEB
$fmswebResolve = Resolve-DnsName fmsweb.fms.army.mil

if($fmswebResolve.name -like "*aaghqda.army.mil.apps.gcds.disa.mil*" -and "*aaghqda.army.mil.edgekey.akamai.csd.disa.mil*")
{
Write-Host -ForegroundColor Green "DNS Alias test for FMSWEB Succeeded"
}
else
{
Write-Host -ForegroundColor Red "DNS Alias test for FMSWEB Failed"
}

Write-Host -ForegroundColor Yellow "Pause for manual check of Websites, if above DNS resolution all succeeds then tests AVD-70 - AVD-73 pass, press enter after tests are complete"
Read-Host #press button

#Bastion Test AVD-79
Write-Host -ForegroundColor Yellow "Pause for manual check of Bation test AVD-79, press enter after tests are complete"
Read-Host #press button

#Fiery Printer Driver AVD-80
Start-Process "https://print.army.mil"
Write-Host -ForegroundColor Cyan "Starting Fiery Printer Driver test AVD-80"

Write-Host -ForegroundColor Yellow "Pause for manual addition of printer in Printer Logic for test AVD-80, press enter after this step has been complete" 
Read-Host #press button

Write-Host -ForegroundColor Cyan "Getting Run keys to check for FPDU Key"
$runKey = Get-Item HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run
if($runKey.Property -eq 'FDPU'){
Write-Host -ForegroundColor -Green 'FDPU Key Exists'
}
else
{
Write-Host -ForegroundColor Green 'FDPU Key Does Not Exist'
}

Get-Item HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run
Write-Host -ForegroundColor Yellow "Pause for manual check of FPDU Key test AVD-80, press enter after tests are complete"
Read-Host #press button

#Certificate Test AVD-81
Write-Host -ForegroundColor Cyan "Gathering DoD and ECMA-WCF-TRUST Trusted Root Certs AND DoD Intermediate Certs"
Start-Sleep -Seconds 2

Get-ChildItem -Path Cert:\CurrentUser\CA | Where-Object { $_.Issuer -like "*DOD*" -or $_.Issuer -like "*ECMA*"} | Sort-Object NotAfter |
ForEach-Object {
$status = "Not Expired"
$color = "Green"

if ($_.NotAfter -lt (Get-Date)) {
$status = "Expired"
$color = "Red"
}

Write-Host ("Issuer: " + $_.Issuer).PadRight(100) -NoNewline
Write-Host ("Expires: " + $_.NotAfter).PadRight(30) -NoNewline
Write-Host ("Status: " + $Status) -ForegroundColor $color

}

Get-ChildItem -Path Cert:\CurrentUser\Root | Where-Object { $_.Issuer -like "*DOD*" -or $_.Issuer -like "*ECMA*"} | select-object issuer, notafter | Sort-Object NotAfter |
ForEach-Object {
$status = "Not Expired"
$color = "Green"

if ($_.NotAfter -lt (Get-Date)) {
$status = "Expired"
$color = "Red"
}

Write-Host ("Issuer: " + $_.Issuer).PadRight(100) -NoNewline
Write-Host ("Expires: " + $_.NotAfter).PadRight(30) -NoNewline
Write-Host ("Status: " + $Status) -ForegroundColor $color

}

Write-Host -ForegroundColor Yellow "Pause for manual check of Root cert expiration. Use above results to check for expired certs for AVD-81 test then press enter"
Read-Host #press button


#Disable Ipv6 Test AVD-82
Write-Host -ForegroundColor Cyan "Ipv6 test AVD-82"
$ipv6 = Get-NetAdapterBinding -ComponentID ms_tcpip6 #| select Enabled
Write-Host -ForegroundColor Green "Use below results for AVD-82 test, this test should show Enabled as FALSE"
$ipv6 | ft -AutoSize

Write-Host -ForegroundColor Yellow "Pause for manual check of Ipv6 setting test AVD-82, press enter after tests are complete"
Read-Host #press button

Write-Host -ForegroundColor Green "UAT Complete"



