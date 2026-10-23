#//check for fixed data drive encryption. Decrypt if they are encrypted;
write-host "checking on Bitlocker encryption status for fixed drives, and decrypting if encrypted..."
ForEach ($volume in Get-BitLockerVolume){
    
    If($volume.EncryptionPercentage -gt 0){

        Disable-BitLocker -MountPoint $volume.mountpoint

        do
        {
            $P = (Get-BitLockerVolume $volume.MountPoint).EncryptionPercentage
            Write-output "BitLocker Volume Percentage for Drive $($volume.MountPoint) $P%. Please wait while decryption completes..."
            sleep 4
        }
        Until ((Get-BitLockerVolume $volume.MountPoint).VolumeStatus -eq "FullyDecrypted" `
                -and (Get-BitLockerVolume $volume.MountPoint).ProtectionStatus -eq `
                "Off" -and (Get-BitLockerVolume $volume.MountPoint).EncryptionPercentage -eq "0")
    }Else{

        Write-Output "$($volume.MountPoint) is not BitLocker encrypted."
    }
    
}

#//check status of BDESVD (Bitlocker Drive Encryption) Service. Ensure it is stopped and disabled;
$service = Get-Service | ? {$_.name -eq "BDESVC"}
Write-Output "Checking on status of Bitlocker Drive Encryption service..."

if ($service.Status -eq "Running"){
    Write-Output "Bitlocker Drive Encryption service is running, stopping it..."
    $cmd = "sc.exe stop $($service.name)"
    Invoke-Expression -Command $cmd
}else{
    Write-Output "Bitlocker Drive Encryption service in stopped state."
}

If ($service.StartType -ne "Disabled"){
    Write-Output "Bitlocker Drive Encryption service is not disabled, disabling it..."
    $cmd = "sc.exe config $($service.name) start=disabled"
    Invoke-Expression -Command $cmd
}else{
    Write-Output "Bitlocker Drive Encryption service in disabled state."
}