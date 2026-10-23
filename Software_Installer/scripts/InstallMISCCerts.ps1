$workDir    = "C:\temp\bootstrap\InstallMISCCerts"
$useZIPfile = $true

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

$InstallerPath = Get-childitem -path . -File

If($useZIPfile -eq $TRUE){
    Write-Output "Extracting contents of $($ZIPFile) to $($workDir)..."
    $ZIPFile = $InstallerPath | Where-Object {$_.Extension -match ".zip"}
    Expand-Archive -literalPath $ZIPFile -DestinationPath $workDir -Force
}

Start-Transcript -Path $logPath

$certArray = @("DoD WCF Signing CA")

ForEach($cert in $certArray){
    
    Switch($cert){

        "DoD WCF Signing CA"{

            $i = 0
            $today = get-date
            $expiryDate = $today.AddDays(30)
            $rootURL  = "https://crl.gds.disa.mil/getsign?"
            $downloadDir = "C:\Windows\temp"
            $bolUPDATE = $FALSE

            Do{

                $i++
                $DNSName = "DoD WCF Signing CA $($i)"

                $certInfo = gci Cert:\LocalMachine\CA | select NotAfter, DNSNameList, Subject | ? {@($_.DnsNameList | ForEach-Object{ $_.unicode}) -contains $DNSName}
                
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
                    $output  = $downloadDir + "\" + $newCertDNSName + ".cer"
                    Invoke-WebRequest -Uri $fullURL -OutFile $output
                    
                    Write-Output "importing $($certInfo.Subject)..."
                    $doImport = Import-Certificate -FilePath $output -CertStoreLocation Cert:\LocalMachine\CA

                    If($doImport.Thumbprint){
                        Write-Output "successfully imported $($certInfo.Subject)."

                        If($certInfo.subject){
                            Write-Output "Removing older version of $($certInfo.subject) that is set to expire..."
                            gci Cert:\LocalMachine\CA | ? {$_.Subject -eq "$($certInfo.Subject)"} | Remove-Item -Force
                        }

                    }
                    
                }

            }until ($i -eq 10)

        };
        
    }

}

#//parse through cert folders in $workdir and add to cert stores based on folder name;
$certFldrs = get-childitem -path "$($workDir)" -Directory

ForEach($fldr in $certFldrs){
    $listOfCerts = get-childitem -path $fldr.FullName -File

    ForEach($cert in $listOfCerts){
        Import-Certificate -FilePath $cert.FullName -CertStoreLocation Cert:\LocalMachine\$fldr -Confirm:$false
    }
    
}

Stop-Transcript
EXIT 0