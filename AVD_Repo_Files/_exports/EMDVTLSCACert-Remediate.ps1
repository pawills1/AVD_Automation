﻿Function Import-Certs{

    param(
        [Parameter(Mandatory=$true)]
        $base64TXT,
        [Parameter(Mandatory=$true)]
        $certStoreName,
        [Parameter(Mandatory=$true)]
        $certLocation
    )
    
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("$($certStoreName)","$($certLocation)")
    $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)

    $certInBytes = [System.Convert]::FromBase64String($base64TXT)
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(,$certInBytes)

    try{
        $store.add($cert)
    }catch{
        $_
    }finally{
        $store.Close()
    }

}#//end function

#############################################################################

$cert = @"
MIID4jCCAsqgAwIBAgIRAPv3cXVNaodgJjs+EVQyWT8wDQYJKoZIhvcNAQELBQAw
ezELMAkGA1UEBhMCR0IxGzAZBgNVBAgMEkdyZWF0ZXIgTWFuY2hlc3RlcjEQMA4G
A1UEBwwHU2FsZm9yZDEaMBgGA1UECgwRQ29tb2RvIENBIExpbWl0ZWQxITAfBgNV
BAMMGEFBQSBDZXJ0aWZpY2F0ZSBTZXJ2aWNlczAeFw0yNTA2MTgwMDAwMDBaFw0y
ODEyMzEyMzU5NTlaMFMxCzAJBgNVBAYTAklOMSUwIwYDVQQKDBxlTXVkaHJhIFRl
Y2hub2xvZ2llcyBMaW1pdGVkMR0wGwYDVQQDDBRFTSBEViBUTFMgQ0EgLSBHM0Et
MTB2MBAGByqGSM49AgEGBSuBBAAiA2IABBFTCyczK1dJiWLL4WF3SUL68GMYeEw8
Md76Yy60An0lPksi0KiaUBTbNfcpupEb+6FrYXdAUTpa8oOeuTT/OnSxXrB/0fBf
ot3HZJ4x1eY2wjf8OkULVUjvkPwYwInmjaOCATUwggExMB8GA1UdIwQYMBaAFKAR
CiM+lvEH7OKvKe+CpX/QMKS0MB0GA1UdDgQWBBSrKJaDaT7ESXH1FIplRyGjcCK2
UzAOBgNVHQ8BAf8EBAMCAQYwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHSUEFjAU
BggrBgEFBQcDAQYIKwYBBQUHAwIwMQYDVR0gBCowKDAIBgZngQwBAgEwDQYLKwYB
BAGDjiEBAAEwDQYLKwYBBAGDjiEBAmQwQwYDVR0fBDwwOjA4oDagNIYyaHR0cDov
L2NybC5jb21vZG9jYS5jb20vQUFBQ2VydGlmaWNhdGVTZXJ2aWNlcy5jcmwwNAYI
KwYBBQUHAQEEKDAmMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5jb21vZG9jYS5j
b20wDQYJKoZIhvcNAQELBQADggEBAIVgEAn/Sh002C8YTPXSFhQW/XX+DfPtO9Nq
VYG4QUzScLtdL1/nF96oV7UWPpPLMNk1ocD8+Yal3V2NpFPaRX8ZN+ra6kOVjvG0
KZpnC6d4R7hFjV6aSzJNqjRLxUz2szE0PGXlxB9KTaws8RhgkTuYExUGJ+3kQ3RW
N9tj2FwnZq1+vJtRZ6ey4yxFYM2EyzthXXXw4nLQwwJAai9/mJLiIFyotGG/SY47
lhYb//JtJsrxyYpboe/JYZC3HpnGvAFcNDFwL8qNJrZBL5M5OwqV/l4dz7593KMC
uav3MFb9BuAjjqnu1WDtcR2aq5LyiZmvCLMBt2dIN/Q1v5+Emyw=
"@.Trim()

Import-Certs -base64TXT $cert -certStoreName 'Disallowed' -certLocation 'LocalMachine'