 $FSLogixDirectoryPath= "$env:ProgramFiles\FSLogix"

 Write-Output "Creating FSLogix redirection XML file under $($FSLogixDirectoryPath)\RedirXMLSourceFolder..."
 If (!(Test-Path -Path "$($FSLogixDirectoryPath)\RedirXMLSourceFolder" -ErrorAction SilentlyContinue)){
    New-Item -ItemType Directory -Path "$($FSLogixDirectoryPath)\RedirXMLSourceFolder" -Force
 }

 New-Item -Path "$($FSLogixDirectoryPath)\RedirXMLSourceFolder\Redirections.xml" -ItemType File -Force -Value '
<?xml version="1.0" encoding="UTF-8"?>

<FrxProfileFolderRedirection ExcludeCommonFolders="0">

<Excludes>
    <Exclude Copy="0">AppData\Local\Google\Chrome\User Data\Default\Cache</Exclude>
    <Exclude Copy="0">AppData\Local\Google\Chrome\User Data\Default\Cached Theme Image</Exclude>
    <Exclude Copy="0">AppData\Local\Google\Chrome\User Data\Default\JumpListIcons</Exclude>
    <Exclude Copy="0">AppData\Local\Google\Chrome\User Data\Default\JumpListIconsOld</Exclude>
    <Exclude Copy="0">AppData\Local\Google\Chrome\User Data\Default\Storage</Exclude>
    <Exclude Copy="0">AppData\Local\Google\Chrome\User Data\Default\Local Storage</Exclude>
    <Exclude Copy="0">AppData\Local\Google\Chrome\User Data\Default\SessionStorage</Exclude>
    <Exclude Copy="0">AppData\Local\Google\Chrome\User Data\Default\Media Cache</Exclude>
    <Exclude Copy="0">AppData\Local\Google\Chrome\User Data\Default\GPUCache</Exclude>
    <Exclude Copy="0">AppData\Local\Google\Chrome\User Data\Default\WebApplications</Exclude>
    <Exclude Copy="0">AppData\Local\Google\Chrome\User Data\Default\SyncData</Exclude>
    <Exclude Copy="0">AppData\Local\Google\Chrome\User Data\Default\SyncDataBackup</Exclude>
    <Exclude Copy="0">AppData\Local\Google\Chrome\User Data\Default\Pepper Data\Shockwave Flash\CacheWriteableAdobeRoot</Exclude>
    <Exclude Copy="0">AppData\Local\Google\Chrome\User Data\WidevineCDM</Exclude>
    <Exclude Copy="0">AppData\Local\Google\Chrome\User Data\EVWhitelist</Exclude>
    <Exclude Copy="0">AppData\Local\Google\Chrome\User Data\pnacl</Exclude>
    <Exclude Copy="0">AppData\Local\Google\Chrome\User Data\recovery</Exclude>
    <Exclude Copy="0">AppData\Local\Google\Chrome\User Data\ShaderCache</Exclude>
    <Exclude Copy="0">AppData\Local\Google\Chrome\User Data\SwReporter</Exclude>
    <Exclude Copy="0">AppData\Local\Google\Chrome\User Data\SwiftShader</Exclude>
    <Exclude Copy="0">AppData\Local\Google\Chrome\User Data\PepperFlash</Exclude>
    <Exclude Copy="0">AppData\Local\Microsoft\Edge\User Data\Default\Cache</Exclude>
    <Exclude Copy="0">AppData\Local\Packages\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\AC\MicrosoftEdge\Cache</Exclude>
    <Exclude Copy="0">AppData\Local\BraveSoftware\Brave-Browser\User Data\Default\Cache</Exclude>
    <Exclude Copy="0">AppData\Local\Microsoft\Windows\WER</Exclude>
    <Exclude Copy="0">AppData\Local\Microsoft\Terminal Server Client\Cache</Exclude>
    <Exclude Copy="0">AppData\Roaming\Downloaded Installations</Exclude>
    <Exclude Copy="0">AppData\Local\Downloaded Installations</Exclude>
    <Exclude Copy="0">AppData\Local\Microsoft\Office\16.0\Lync\Tracing</Exclude>
    <Exclude Copy="0">AppData\Local\Microsoft\MSOIdentityCRL\Tracing</Exclude>
    <Exclude Copy="0">AppData\Local\Microsoft\OneNote\16.0\Backup</Exclude>
    <Exclude Copy="0">AppData\Local\CrashDumps</Exclude>
    <Exclude Copy="0">AppData\Local\SquirrelTemp</Exclude>
    <Exclude Copy="0">AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Logs</Exclude>
    <Exclude Copy="0">AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\PerfLog</Exclude>
    <Exclude Copy="0">AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\WV2Profile_tfw\GPUCache</Exclude>
    <Exclude Copy="0">AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\WV2Profile_tfw\WebStorage</Exclude>
    <Exclude Copy="0">AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\WV2Profile_dod\GPUCache</Exclude>
    <Exclude Copy="0">AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\WV2Profile_dod\WebStorage</Exclude>
</Excludes>

<Includes>
    <Include Copy="3">AppData\LocalLow\Sun\Java\Deployment\security</Include>
</Includes>

</FrxProfileFolderRedirection>'

Set-ItemProperty -Path "HKLM:\Software\FSLogix\Profiles" -Name "RedirXMLSourceFolder" -Value "$($FSLogixDirectoryPath)\RedirXMLSourceFolder" -Force