$workDir    = "C:\temp\bootstrap\DoDWCFSigningCerts"

$useZIPfile = $false

#//if $workdir specified, sets that as working location (if it exists). Otherwise, assumes script location is working directory.
If($workDir){

    If(Test-Path -Path "$($workdir)"){
        $origLocation = (Get-Location).path
        Set-Location -Path $($workDir)
    }else{
        If ($useZIPfile -eq $false){
            Write-Output "The working path $($workDir) does not exit. Creating it now"
            New-Item -ItemType Directory -Path $workDir -Force
            Set-Location -Path $($workDir)
        }else{
            Write-Warning "The working path $($workDir) does not exit. Script will exit."
            exit
        }
    }

}

#determine installation log path;
If(!(test-path -Path "$env:SystemDrive\Program Files\AVDAppInstallLogs")){
    New-Item -ItemType Directory -Path "$env:SystemDrive\Program Files\AVDAppInstallLogs" -Force
}

$logFolder = "$env:SystemDrive\Program Files\AVDAppInstallLogs"
$appNameForLog = $($workDir).split('\')[3]
$LogPath = "$($logFolder)\$($appNameForLog)_AVD_Script.log"

Start-Transcript -Path $logPath

$i = 0
$today      = get-date
$expiryDate = $today.AddDays(30)
$rootURL    = "https://crl.gds.disa.mil/getsign?"
$bolUPDATE  = $FALSE

Do{

    $i++
    $DNSName = "DoD WCF Signing CA $($i)"

    $certInfo = gci Cert:\LocalMachine\CA | select NotAfter, DNSNameList, Subject | ? {($_.DNSNameList -eq "$($DNSName)")}
    
    If(!($certInfo)){
        Write-Output "$($DNSName) certificate is missing..."
        $bolUPDATE = $TRUE
    }ElseIf ($certInfo.NotAfter -le $expiryDate){
        Write-Output "$($DNSName) certificate is set to expire on $($certInfo.NotAfter)..."
        $bolUPDATE = $TRUE
    }

    If($bolUpdate -eq $TRUE){
        $newCertDNSName = $DNSName.TOUpper() -replace (" ", "+")

        $fullURL = $rootURL + $newCertDNSName
        $output  = $workdir + "\" + $newCertDNSName + ".cer"
        Invoke-WebRequest -Uri $fullURL -OutFile $output
        
        Write-Output "importing $($certInfo.Subject)..."
        $doImport = Import-Certificate -FilePath $output -CertStoreLocation Cert:\LocalMachine\CA

        If($doImport.Thumbprint){
            Write-Output "successfully imported $($certInfo.Subject)..."

            If($certInfo.subject){
                Write-Output "Removing older version of $($certInfo.subject) that is set to expire..."
                gci Cert:\LocalMachine\CA | ? {$_.Subject -eq "$($certInfo.Subject)"} | Remove-Item -Force
            }

        }
        
    }

    $bolUPDATE = $FALSE

}until ($i -eq 10)

Stop-Transcript
EXIT 0