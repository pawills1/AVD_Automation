$assetsStorageAccount = $env:PKR_VAR_asset_storage_account | ConvertFrom-Json -AsHashtable
$ctx = (Get-AzStorageAccount -ResourceGroupName $assetsStorageAccount.resource_group_name -Name $assetsStorageAccount.name).Context
$allBlobs = Get-AzStorageBlob -Container $assetsStorageAccount.rd_agents_container_name -Context $ctx

$PinnedFiles = @(
        $env:PKR_VAR_rd_agent_installer,
        $env:PKR_VAR_rd_agent_bootloader
    )

Write-Host "Syncing pin status for: '$($PinnedFiles -join ', ')'"

foreach ($blob in $allBlobs) {
    $blobClient = $blob.ICloudBlob
    $blobClient.FetchAttributes()
    $isPinned = $blobClient.Metadata.ContainsKey("pinned")

    if ($PinnedFiles -contains $blob.Name) {
        if (-not $isPinned) {
            Write-Host "Pinning: '$($blob.Name)'"
            $blobClient.Metadata["pinned"] = "true"
            $blobClient.SetMetadata()
        } else {
            Write-Host "Already pinned: '$($blob.Name)'"
        }
    }

    elseif ($isPinned) {
        Write-Host "Unpinning blob: '$($blob.Name)'"
        $blobClient.Metadata.Remove("pinned") | Out-Null
        $blobClient.SetMetadata()
    }
}

$finalPinned = $allBlobs | Where-Object { $_.ICloudBlob.Metadata.ContainsKey("pinned") }
Write-Host "Verification: '$($finalPinned.Count)' blob(s) currently pinned."