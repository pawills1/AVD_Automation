[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [String]
    $BranchName
)

# All errors are terminating errors
$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

Write-Output "Getting gallery image definition for branch '$BranchName'"

$GalleryImageDefinitions = $env:PKR_VAR_gallery_image_definitions | ConvertFrom-Json -AsHashtable -Depth 100
$GalleryImageDefinition = switch -Regex ($BranchName) {
    "^(release-slice-a|release-slice-a-.+)$" {
        $GalleryImageDefinitions.GetEnumerator().where({ $PSItem.name -ilike "*-slice-a" })
    }
    "^(release-slice-b|release-slice-b-.+)$" {
        $GalleryImageDefinitions.GetEnumerator().where({ $PSItem.Name -ilike "*-slice-b" })
    }
    "^(release-base-.+)$" {
        $GalleryImageDefinitions.GetEnumerator().where({ $PSItem.Name -ilike "*-base" })
    }
    "^(release-updates-.+)$" {
        $GalleryImageDefinitions.GetEnumerator().where({ $PSItem.Name -ilike "*-updates" })
    }
    "^(release-general-.+)$" {
        $GalleryImageDefinitions.GetEnumerator().where({ $PSItem.Name -ilike "*-general" })
    }
    default {
        $GalleryImageDefinitions.GetEnumerator().where({
                (-not ($PSItem.Name -ilike "*-slice-a")) -and
                (-not ($PSItem.Name -ilike "*-slice-b")) -and
                (-not ($PSItem.Name -ilike "*-base")) -and
                (-not ($PSItem.Name -ilike "*-updates")) -and
                (-not ($PSItem.Name -ilike "*-general"))
            }
        )
    }
}

if ($GalleryImageDefinition.Count -eq 0) {
    throw "Failed to find a gallery image definition for branch '$BranchName'"
} elseif ($GalleryImageDefinition.Count -gt 1) {
    $GalleryImageDefinition
    throw "Found more than 1 gallery image definition for branch '$BranchName'"
}

Write-Output "Found gallery image definition '$($GalleryImageDefinition.Name)' for branch '$BranchName'"
Write-Output "NAME=$($GalleryImageDefinition.Name)" >> $env:GITHUB_OUTPUT
