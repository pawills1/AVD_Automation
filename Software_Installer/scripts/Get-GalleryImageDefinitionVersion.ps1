[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [String]
    $BuildTag,

    [Parameter(Mandatory)]
    [String]
    $GalleryImageDefinitionName
)

Import-Module -Name ./ArmyEitaasDevops/ArmyEitaasDevops.psm1

$GalleryImageDefinitions = $env:PKR_VAR_gallery_image_definitions | ConvertFrom-Json -AsHashtable

Write-Output "Using image definition '$GalleryImageDefinitionName' to get image version"
$GalleryImageDefinition = $GalleryImageDefinitions.$GalleryImageDefinitionName
$Parameters = @{
    GalleryImageDefinitionName = $GalleryImageDefinitionName
    GalleryName                = $GalleryImageDefinition.gallery_name
    ResourceGroupName          = $GalleryImageDefinition.resource_group_name
    IncrementPosition          = $null
}

$GalleryImageVersions = Get-AzGalleryImageVersion `
    -ResourceGroupName $GalleryImageDefinition.resource_group_name `
    -GalleryName $GalleryImageDefinition.gallery_name `
    -GalleryImageDefinitionName $GalleryImageDefinitionName `
    -ErrorAction SilentlyContinue

Write-Output "Image is being build, getting next image version"
if (($GalleryImageVersions.Count -gt 0) -or ($null -ne $GalleryImageVersions)) {
    $LatestGalleryImageVersion = $GalleryImageVersions | Sort-Object -Property Name | Select-Object -Last 1
    $LatestBuildTag = $LatestGalleryImageVersion.Tags["build_tag"]
    if ($null -eq $LatestBuildTag) {
        throw "Gallery image definition '$GalleryImageDefinitionName' exists but build tag is null"
    }
}

if ($LatestBuildTag -ne $BuildTag) {
    Write-Output "Latest build tag '$LatestBuildTag' doesn't match current build tag '$BuildTag', incrementing minor position"
    $Parameters.IncrementPosition = "Minor"
} else {
    Write-Output "Latest build tag '$LatestBuildTag' matches current build tag '$BuildTag', incrementing build position"
    $Parameters.IncrementPosition = "Build"
}

$NextVersion = Get-GalleryImageSemVersion @Parameters -Verbose
Write-Output "Next image version is '$NextVersion'"
Write-Output "VERSION=$NextVersion" >> $env:GITHUB_OUTPUT
Write-Output "EXCLUDE=false" >> $env:GITHUB_OUTPUT
