function Resolve-UatTemplate {
    param([string]$UatDirectory, [string]$InputFile)

    if ($InputFile) {
        $file = Get-Item -LiteralPath $InputFile -ErrorAction Stop
        if ($file.PSIsContainer -or $file.Extension -ne '.xlsx') {
            throw "Input checklist must be an .xlsx file: $InputFile"
        }
        return $file.FullName
    }

    $files = @(Get-ChildItem -LiteralPath $UatDirectory -Filter 'AVD Test Plan*.xlsx' -File |
        Where-Object { $_.Name -notlike '~$*' } |
        Sort-Object LastWriteTime -Descending)
    if ($files.Count -eq 0) {
        throw "No Excel workbook matching 'AVD Test Plan*.xlsx' was found in '$UatDirectory'. Deploy the source checklist to this directory."
    }
    return $files[0].FullName
}

function Resolve-UatAutomatedWorkbook {
    param([string]$UatDirectory)

    $directories = @(Get-ChildItem -LiteralPath $UatDirectory -Directory -Filter 'Automated_Results_*' |
        Where-Object { $_.Name -match '^Automated_Results_\d{8}-\d{6}(?:-\d{3})?$' } |
        Sort-Object Name -Descending)
    if ($directories.Count -eq 0) {
        throw "No Automated_Results_<timestamp> folder was found in '$UatDirectory'. Run the automated check first."
    }
    $latest = $directories[0].FullName
    $workbooks = @(Get-ChildItem -LiteralPath $latest -File -Filter 'AVD Checklist - Completed - *.xlsx' |
        Sort-Object Name -Descending)
    if ($workbooks.Count -eq 0) {
        throw "The newest automated-results folder '$latest' has no completed .xlsx checklist. Complete the automated run before starting manual checks."
    }
    return $workbooks[0].FullName
}

function New-UatResultsDirectory {
    param([string]$RootDirectory)

    $path = Join-Path $RootDirectory ("Automated_Results_{0}" -f (Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
    return (New-Item -ItemType Directory -Path $path -ErrorAction Stop).FullName
}

# SIG # Begin signature block
# MIIHMAYJKoZIhvcNAQcCoIIHITCCBx0CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDflRNpm5iS6N3R
# cXaBwWaHwdm7Zz7xRHWfVkjfpqJBfaCCBBwwggQYMIICgKADAgECAhARPahbg4yR
# u0x0RWe84V06MA0GCSqGSIb3DQEBCwUAMCQxIjAgBgNVBAMMGUFWRCBVQVQgVGVz
# dCBDb2RlIFNpZ25pbmcwHhcNMjYwOTIzMTk0MjUxWhcNMzEwOTIzMTk1MjUwWjAk
# MSIwIAYDVQQDDBlBVkQgVUFUIFRlc3QgQ29kZSBTaWduaW5nMIIBojANBgkqhkiG
# 9w0BAQEFAAOCAY8AMIIBigKCAYEAr3U0akZAjTYFpKBuGxG/XOZrFq1TiCKvagLt
# k5EbZnFvlNami9JqgzikK4A/wFJxQBdhaffseFIxS0T7HearbpAtDuQTVdkhcC03
# uBMs68EyzlggBzT2hOH7v9ebpQ3l0iz23ubNUXAtevgda1KVHfpNdTo1I1kSE6vt
# dWzVx4XyrJDC/2kK7yMpFZAw1B0VqrNOt2EQDrBAYofgIQxB6QEmp7b7O33qlIBq
# hGJFziEsTqrGbr0oCQKiXblwmUS1qKl+hMb93xkZ29WiC3PCOpnW+jwEN5zfGRqH
# 1Q0j1eVEoFTNWuEJ81N7HTi7tNAC71gO/4PNQJ0562UFkovR9fUX00UUBjKCBSqy
# nqYIKKqSErrOAFYepr4bOpheVPNsWxxwjxLWfBze92kX9IgNN3otMqTvVFY2nYXl
# hhbN217V3yGoqVmaA6a+AIiBYhXt3WegcDtP02MX/mZvbZpeHC6QpqlKHZVxgdth
# cN/5evLCaoGxxBO85VJwFM/JKbUpAgMBAAGjRjBEMA4GA1UdDwEB/wQEAwIHgDAT
# BgNVHSUEDDAKBggrBgEFBQcDAzAdBgNVHQ4EFgQUkggBwTJ4XHabRzNkW8BAJOxx
# UXAwDQYJKoZIhvcNAQELBQADggGBAAZBilHu50OctPP6mSsw/AxFk/gdNzSr/1dC
# Hi24+oU6t6bcZcRp1JJXKbipWLHfR1z3QMjuOzbwVEOlBfq0rQRygVVrKkH21MXs
# ubIjl1AWXdkwrnbJl0TpiCBDQBBP+obPS+eJY524ovOBY2tZvI3jgOxIVJDDMRwu
# hB/5BzR/n7V1VUFRzRPeVWTsgjWC+qbkbzrPMgqk1yOinBtFdNRiU67lpMR8OKZx
# UiMVhacc/u8BsEaqFVLJD2jcXe8tnjqYpFHUdb1J6pe0rooMf6oTbQtBtokM5hcx
# rh+kkvuQeY0V340WFGKw1DfTJE9HCnwkGPR94KgVauYWmF3dF3qDlvAT6mUqZxT2
# LyR98EnVSlbS+cFWf/2jSEGAZlbv1axCS1JHaypwF3Lo0JUGbzkgE4eupgv65w0L
# bzOkrUw/s3ZTVzuzfHW4KFiGOhzwqO26lsviuWZtefpsPB7mhp6B65wo7Bt+wTrx
# Nbarmp4AGEZFrhszz4hGkaX1swYgaDGCAmowggJmAgEBMDgwJDEiMCAGA1UEAwwZ
# QVZEIFVBVCBUZXN0IENvZGUgU2lnbmluZwIQET2oW4OMkbtMdEVnvOFdOjANBglg
# hkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3
# DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEV
# MC8GCSqGSIb3DQEJBDEiBCBz/9LSrezoPUESbIPyBSMIXPCFog+eskGw7hBW5DTQ
# BDANBgkqhkiG9w0BAQEFAASCAYA5EW9pSwb9fUwba1cW6DIy2/Zsbu4gez+aQNre
# 7+rf4r9V9nC5omoss0nhrRFmjG7wIwk4yRBligJAUVAxm6b0Vmw6ZTAN1pF5NMUt
# PzJA/oPgMbH+ArRe3Fp6m+pY997rzdfOHAvKjY0vhFyJ6mmreLAhXSSPOScqGc/p
# livpvSYif4YwxJYRgLrxb5fb4/uG+PQVsdGzdyqSDjmAxyS+4KXU7TOg1voJ2Ws6
# 21vs9xlyFGqHmO8yq4p4n2f+TpgUJ4x0nSCdXlg7KpX50qJ4pEuSeVu1mdHHpPed
# dnq0h9ijvTMYV5nI8sckd7IMhs8M6j+pXnu5ewmAR0fsLIysnFKBROjW0GTMNvYS
# 5GCWQhaogHD6dfc1H7Do0LHmcg6rnG4Wu2xaIHvY32FgX9TgbFA53zdbzrvVo9dX
# uQEik1Lv7L+yhJrBrY7ir6h+j5q+viUaopfecKwGa06458BwBXFce9GDl3msc1Q0
# EzQtPNjEBMNh4uyzVMS0j504www=
# SIG # End signature block
