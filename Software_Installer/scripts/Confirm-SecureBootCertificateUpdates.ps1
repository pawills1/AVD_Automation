$loopCount = 0
Do{

    Start-Sleep -Seconds 5

    If (Get-WinEvent -ProviderName "Microsoft-Windows-TPM-WMI" | ? {$_.Id -eq '1808'}){
        break
    }else{
        Write-Output "Awaiting completion of Secure Boot Certificate Updates. Please wait..."
        Start-Sleep -Seconds 25
        $loopCount++
    }

}until((Get-WinEvent -ProviderName "Microsoft-Windows-TPM-WMI" | ? {$_.Id -eq '1808'}) -OR $loopCount -eq 11)

If (Get-WinEvent -ProviderName "Microsoft-Windows-TPM-WMI" | ? {$_.Id -eq '1808'}){
    Write-Output "Secure boot certificate updates have been successfully applied to this system. Continuing on with image build."
    EXIT 0
}Else{
    Write-warning "Secure boot certificate updates have NOT been successfully applied to this system. Please investigate VM for further information."
    EXIT 1
}