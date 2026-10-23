$arrCertSubjects = @("CN=EM DV TLS CA - G3-1")

$arrHTCertInfo = @()
$arrHTCertInfo = $arrHTCertInfo + [pscustomobject]@{Subject="CN=EM DV TLS CA - G3-1";Path="Cert:\LocalMachine\Disallowed"}

$COMPLIANT = $TRUE

ForEach($certInfo in $arrHTCertInfo){

    $getCert = Get-ChildItem $certInfo.Path | ? {$_.subject -like "$($certInfo.Subject)*"}

    If (!($getCert)){
        $COMPLIANT = $false
        break
    }

}