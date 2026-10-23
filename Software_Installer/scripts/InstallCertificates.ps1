$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

function Get-TimeStamp { return "[{0:MM/dd/yy} {0:HH:mm:ss} UTC]" -f (Get-Date).ToUniversalTime() }

##########################
## Install Certificates ##
##########################

Write-Output "$(Get-TimeStamp) - Starting Script - Install Certificates"

$CertDirectoryPath = "C:\downloaded-certs\"
if (-not (Test-Path $CertDirectoryPath)) {
    throw "Failed to find extracted certificates temporary directory."
}

$Certificates = Get-ChildItem -Path $CertDirectoryPath -File -Recurse
if ($Certificates.Count -eq 0) {
    throw "Found 0 certificates within extracted certificates temporary directory."
}

Write-Output "$(Get-TimeStamp) - Found $($Certificates.Count) certificates to install."
foreach ($Certificate in $Certificates) {
    $ImportedCertificate = switch -Wildcard ($Certificate.Directory.Name) {
        "root" {
            $StoreLocation = "Microsoft Local Computer Root Store"
            Import-Certificate -FilePath $Certificate.FullName -CertStoreLocation Cert:\LocalMachine\Root
        }

        "intermediate" {
            $StoreLocation = "Microsoft Local Computer Intermediate Store"
            Import-Certificate -FilePath $Certificate.FullName -CertStoreLocation Cert:\LocalMachine\CA
        }

        "untrusted" {
            $StoreLocation = "Microsoft Local Computer Untrusted Store"
            Import-Certificate -FilePath $Certificate.FullName -CertStoreLocation Cert:\LocalMachine\Disallowed
        }
    }
    Write-Output "$(Get-TimeStamp) - Successfully imported certificate (Subject: $($ImportedCertificate.Subject), Issuer: $($ImportedCertificate.Issuer)) into $StoreLocation."
}

Write-Output "$(Get-TimeStamp) - Completed Script - Install Certificates"
