$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

function Get-TimeStamp { return "[{0:MM/dd/yy} {0:HH:mm:ss} UTC]" -f (Get-Date).ToUniversalTime() }

############################
## Download Storage Blobs ##
############################

Write-Output "$(Get-TimeStamp) - Starting Script - Download Storage Blobs"

$TmpDirectoryPath = "C:\Temp"
$BootstrapDirectoryPath = Join-Path -Path $TmpDirectoryPath -ChildPath "bootstrap"
if (-not (Test-Path $BootstrapDirectoryPath)) {
    New-Item -Path $BootstrapDirectoryPath -ItemType Directory -Force | Out-Null
}

# Disable automatic saving of Azure context and credentials to local disk.
Disable-AzContextAutosave -Scope Process | Out-Null

$SpCredentials = (New-Object System.Management.Automation.PSCredential ($env:SpClientId, (ConvertTo-SecureString -String $env:SpClientSecret -AsPlainText -Force)))
Connect-AzAccount -Credential $SpCredentials -SubscriptionId $env:SubscriptionId -Tenant $env:TenantId -Environment "AzureUSGovernment" -ServicePrincipal

$StorageAccountKey = (Get-AzStorageAccountKey `
        -ResourceGroupName $env:StorageAccountRgName `
        -Name $env:StorageAccountName).Where({ $PSItem.KeyName -eq "Key1" }).Value

$StorageContext = New-AzStorageContext `
    -StorageAccountName $env:StorageAccountName `
    -StorageAccountKey $StorageAccountKey

Write-Output "$(Get-TimeStamp) - Attempting to download blobs from storage account '$env:StorageAccountName' and container '$env:StorageContainerName'"
$StorageBlobNames = $env:StorageBlobNamesCsv.Split(",").ForEach({ $PSItem.Trim() })

$index = @()
foreach($i in $StorageBlobNames){
    if($i -match "_AVD.zip"){
        $index += $([array]::IndexOf($StorageBlobNames,$i))
    }
}

$avdZipElements = $($StorageBlobNames[$index[0]..($StorageBlobNames.count-1)])

foreach ($BlobName in $StorageBlobNames) {

    $trimmedBlobName = $null
    $BlobDestination = Join-Path -Path $BootstrapDirectoryPath -ChildPath $BlobName
    if ($blobName -match "_AVD.zip"){
        $trimmedBlobName = $blobName.Trim("_AVD.zip").Replace("-Z","-Zip").Replace("dobe","Adobe").Replace("Deskto","Desktop").Replace("SCConfig","DSCConfig").Replace("XStudio","DAXStudio").Replace("TSLite","DTSLite").Replace("SelfHel","SelfHelp") #Don't look at this
        Write-Output "Trimmed blob name : $($trimmedBlobName)"
    }
    $updatedPath = Join-Path -Path $BootstrapDirectoryPath -ChildPath $trimmedBlobName

    try {
        if ($avdZipElements.Contains($blobName)){

            $finalPath = "$updatedPath\$BlobName"
            if (-not (Test-Path $updatedPath)) {
                New-Item -Path $updatedPath -ItemType Directory -Force | Out-Null
            }

            Write-Output "$(Get-TimeStamp) - Downloading blob '$BlobName' to path '$($finalPath)'"
            Get-AzStorageBlobContent `
                -Context $StorageContext `
                -Container $env:StorageContainerName `
                -Blob $BlobName `
                -Destination $updatedPath | Out-Null

            if (-not (Test-Path $finalPath)) {
                throw "Failed to find downloaded blob '$BlobName' '$finalPath'"
            }
        }
        if (-not($avdZipElements.Contains($blobName))){
            Write-Output "$(Get-TimeStamp) - Downloading blob '$BlobName' to path '$BlobDestination'"
            Get-AzStorageBlobContent `
                -Context $StorageContext `
                -Container $env:StorageContainerName `
                -Blob $BlobName `
                -Destination $BlobDestination | Out-Null

            if (-not (Test-Path $BlobDestination)) {
                throw "Failed to find downloaded blob '$BlobName' '$BlobDestination'"
            }
        }
    }
    catch {
        Write-Warning "Failed to download blob '$BlobName' from storage account '$env:StorageAccountName' and container '$env:StorageContainerName'"
        #throw $PSItem
    }
}

Write-Output "$(Get-TimeStamp) - Listing content of downloaded blobs in directory '$BootstrapDirectoryPath'"
Get-ChildItem -Path $BootstrapDirectoryPath -Recurse

Write-Output "$(Get-TimeStamp) - Completed Script - Download Storage Blobs"
