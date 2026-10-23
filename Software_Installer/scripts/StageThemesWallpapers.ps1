$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

function Get-TimeStamp { return "[{0:MM/dd/yy} {0:HH:mm:ss} UTC]" -f (Get-Date).ToUniversalTime() }

###############################
## Stage Themes & Wallpapers ##
###############################

Write-Output "$(Get-TimeStamp) - Starting Script - Stage Themes & Wallpapers"

$TmpDirectoryPath = "C:\Temp"
$BootstrapDirectoryPath = Join-Path -Path $TmpDirectoryPath -ChildPath "bootstrap"
if (-not (Test-Path $BootstrapDirectoryPath)) {
    New-Item -Path $BootstrapDirectoryPath -ItemType Directory -Force | Out-Null
}

$ThemesWallpapersDirectoryPath = Join-Path -Path $BootstrapDirectoryPath -ChildPath "themes-wallpapers"
$ThemesDirectoryPath = "C:\Windows\Resources\Themes"
$WallpapersDirectoryPath = "C:\Windows\Web\Wallpaper\Windows"
foreach ($DirectoryPath in @($ThemesWallpapersDirectoryPath, $ThemesDirectoryPath, $WallpapersDirectoryPath)) {
    Write-Output "$(Get-TimeStamp) - Creating directory '$DirectoryPath'"
    New-Item -Path $DirectoryPath -ItemType Directory -Force | Out-Null
    if (-not (Test-Path $DirectoryPath)) {
        throw "Failed to find directory '$DirectoryPath'"
    }
}

$ThemesZipFilePath = "Army AVD Themes.zip"
$WallpapersZipFilePath = "Army AVD Wallpaper.zip"
foreach ($ZipFileName in @($ThemesZipFilePath, $WallpapersZipFilePath)) {
    $ZipFilePath = Join-Path -Path $BootstrapDirectoryPath -ChildPath $ZipFileName
    if (-not (Test-Path $ZipFilePath)) {
        throw "Failed to find zip file '$ZipFilePath'"
    }

    Write-Output "$(Get-TimeStamp) - Extracting zip file '$ZipFilePath' to directory '$ThemesWallpapersDirectoryPath'"
    Expand-Archive -LiteralPath $ZipFilePath -DestinationPath $ThemesWallpapersDirectoryPath -Force
}

$ThemeFiles = Get-ChildItem -Path $ThemesWallpapersDirectoryPath -Recurse -File -Filter "*.theme"
if (($ThemeFiles.Count -eq 0) -or ($null -eq $ThemeFiles)) {
    throw "Failed to find extracted theme files within directory '$ThemesWallpapersDirectoryPath'"
}

foreach ($File in $ThemeFiles) {
    Write-Output "$(Get-TimeStamp) - Moving theme file '$($File.Name)' to directory '$ThemesDirectoryPath'"
    Move-Item -Path $File.FullName -Destination $ThemesDirectoryPath -Force
}

$WallpaperFiles = Get-ChildItem -Path $ThemesWallpapersDirectoryPath -Recurse -File -Filter "*.jpg"
if (($WallpaperFiles.Count -eq 0) -or ($null -eq $WallpaperFiles)) {
    throw "Failed to find extracted wallpaper files within directory '$ThemesWallpapersDirectoryPath'"
}

foreach ($File in $WallpaperFiles) {
    Write-Output "$(Get-TimeStamp) - Moving theme file '$($File.Name)' to directory '$WallpapersDirectoryPath'"
    Move-Item -Path $File.FullName -Destination $WallpapersDirectoryPath -Force
}

Write-Output "$(Get-TimeStamp) - Completed Script - Stage Themes & Wallpapers"
