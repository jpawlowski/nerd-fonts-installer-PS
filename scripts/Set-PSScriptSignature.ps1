<#PSScriptInfo
.VERSION 1.0.1
.GUID f5fe74f4-b2ef-4f83-8325-15519c9c92fb
.AUTHOR Julian Pawlowski
.COPYRIGHT © 2024 Julian Pawlowski.
.TAGS code-signing powershell powershell-script Windows PSEdition_Core PSEdition_Desktop
.LICENSEURI https://github.com/jpawlowski/nerd-fonts-installer-PS/blob/main/LICENSE.txt
.PROJECTURI https://github.com/jpawlowski/nerd-fonts-installer-PS
.ICONURI https://raw.githubusercontent.com/jpawlowski/nerd-fonts-installer-PS/main/images/nerd-fonts-logo.png
.EXTERNALMODULEDEPENDENCIES
.REQUIREDSCRIPTS
.EXTERNALSCRIPTDEPENDENCIES
.RELEASENOTES
    Version 1.0.0 (2024-09-02)
    - Initial release.
#>

<#
.SYNOPSIS
    Signs a PowerShell script with a specified or selected code signing certificate.

.DESCRIPTION
    This script signs a PowerShell script using a code signing certificate from the current user's certificate store.
    It supports specifying the certificate thumbprint or selecting a certificate from a grid view.
    The script is only supported on Windows.

.PARAMETER ScriptPath
    The path to the PowerShell script that needs to be signed. This parameter is mandatory.

.PARAMETER CertThumbPrint
    The thumbprint of the code signing certificate to be used for signing the script. If not provided, a certificate selection grid view will be displayed.

.EXAMPLE
    .\Set-PSScriptSignature.ps1 -ScriptPath "C:\Scripts\MyScript.ps1"
    Signs the script located at "C:\Scripts\MyScript.ps1" using a selected code signing certificate.

.EXAMPLE
    .\Set-PSScriptSignature.ps1 -ScriptPath "C:\Scripts\MyScript.ps1" -CertThumbPrint "ABC123DEF456..."
    Signs the script located at "C:\Scripts\MyScript.ps1" using the certificate with the specified thumbprint.

.NOTES
    The script uses SHA256 as the hash algorithm and includes the certificate chain except the root certificate.
    A timestamp server is used to timestamp the signature.
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact='High')]
param(
    [Parameter(Mandatory=$true)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$ScriptPath,

    [ValidateNotNullOrEmpty()]
    [string]$CertThumbPrint
)

try {
    if (-not $IsWindows) {
        throw "This script is only supported on Windows."
    }

    if (-not (Get-Command -Name $ScriptPath -ErrorAction SilentlyContinue | Where-Object { $_.CommandType -eq 'ExternalScript' })) {
        throw "$ScriptPath is not a valid PowerShell script."
    }

    $certs = Get-ChildItem -Path Cert:\CurrentUser\My -CodeSigningCert

    if ($certs.Count -eq 0) {
        throw 'No code signing certificate found.'
    }

    if ($CertThumbPrint) {
        $cert = $certs | Where-Object { $_.Thumbprint -eq $CertThumbPrint }
        if (-not $cert) {
            throw "Certificate $CertThumbPrint not found."
        }
    }
    else {
        $cert = $certs | Out-GridView -OutputMode Single -Title 'Select a code signing certificate'
    }

    if (-not $cert) {
        throw 'No certificate selected.'
    }

    if ($PSCmdlet.ShouldProcess("Sign script $ScriptPath with certificate $($cert.Subject)")) {
        $params = @{
            FilePath = $ScriptPath
            Certificate = $cert
            HashAlgorithm = 'SHA256'
            IncludeChain = 'NotRoot'
            TimestampServer = 'http://timestamp.digicert.com'
        }
        Set-AuthenticodeSignature @params
    }

    Write-Host "Script $ScriptPath signed with certificate $($cert.Subject)" -ForegroundColor Green
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
# SIG # Begin signature block
# MIInkQYJKoZIhvcNAQcCoIIngjCCJ34CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAqEqxGv0XPC4aF
# HsBrL+Mc9KUXDEZMKzj4+ANMBifRlaCCIKcwggWNMIIEdaADAgECAhAOmxiO+dAt
# 5+/bUOIIQBhaMA0GCSqGSIb3DQEBDAUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNV
# BAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0yMjA4MDEwMDAwMDBa
# Fw0zMTExMDkyMzU5NTlaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lD
# ZXJ0IFRydXN0ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoC
# ggIBAL/mkHNo3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3E
# MB/zG6Q4FutWxpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKy
# unWZanMylNEQRBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsF
# xl7sWxq868nPzaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU1
# 5zHL2pNe3I6PgNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJB
# MtfbBHMqbpEBfCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObUR
# WBf3JFxGj2T3wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6
# nj3cAORFJYm2mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxB
# YKqxYxhElRp2Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5S
# UUd0viastkF13nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+x
# q4aLT8LWRV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjggE6MIIB
# NjAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qYrhwP
# TzAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzAOBgNVHQ8BAf8EBAMC
# AYYweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdp
# Y2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwRQYDVR0fBD4wPDA6oDigNoY0
# aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENB
# LmNybDARBgNVHSAECjAIMAYGBFUdIAAwDQYJKoZIhvcNAQEMBQADggEBAHCgv0Nc
# Vec4X6CjdBs9thbX979XB72arKGHLOyFXqkauyL4hxppVCLtpIh3bb0aFPQTSnov
# Lbc47/T/gLn4offyct4kvFIDyE7QKt76LVbP+fT3rDB6mouyXtTP0UNEm0Mh65Zy
# oUi0mcudT6cGAxN3J0TU53/oWajwvy8LpunyNDzs9wPHh6jSTEAZNUZqaVSwuKFW
# juyk1T3osdz9HNj0d1pcVIxv76FQPfx2CWiEn2/K2yCNNWAcAgPLILCsWKAOQGPF
# mCLBsln1VWvPJ6tsds5vIy30fnFqI2si/xK4VC0nftg62fC2h5b9W9FcrBjDTZ9z
# twGpn1eqXijiuZQwggauMIIElqADAgECAhAHNje3JFR82Ees/ShmKl5bMA0GCSqG
# SIb3DQEBCwUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMx
# GTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRy
# dXN0ZWQgUm9vdCBHNDAeFw0yMjAzMjMwMDAwMDBaFw0zNzAzMjIyMzU5NTlaMGMx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjE7MDkGA1UEAxMy
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0MDk2IFNIQTI1NiBUaW1lU3RhbXBpbmcg
# Q0EwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDGhjUGSbPBPXJJUVXH
# JQPE8pE3qZdRodbSg9GeTKJtoLDMg/la9hGhRBVCX6SI82j6ffOciQt/nR+eDzMf
# UBMLJnOWbfhXqAJ9/UO0hNoR8XOxs+4rgISKIhjf69o9xBd/qxkrPkLcZ47qUT3w
# 1lbU5ygt69OxtXXnHwZljZQp09nsad/ZkIdGAHvbREGJ3HxqV3rwN3mfXazL6IRk
# tFLydkf3YYMZ3V+0VAshaG43IbtArF+y3kp9zvU5EmfvDqVjbOSmxR3NNg1c1eYb
# qMFkdECnwHLFuk4fsbVYTXn+149zk6wsOeKlSNbwsDETqVcplicu9Yemj052FVUm
# cJgmf6AaRyBD40NjgHt1biclkJg6OBGz9vae5jtb7IHeIhTZgirHkr+g3uM+onP6
# 5x9abJTyUpURK1h0QCirc0PO30qhHGs4xSnzyqqWc0Jon7ZGs506o9UD4L/wojzK
# QtwYSH8UNM/STKvvmz3+DrhkKvp1KCRB7UK/BZxmSVJQ9FHzNklNiyDSLFc1eSuo
# 80VgvCONWPfcYd6T/jnA+bIwpUzX6ZhKWD7TA4j+s4/TXkt2ElGTyYwMO1uKIqjB
# Jgj5FBASA31fI7tk42PgpuE+9sJ0sj8eCXbsq11GdeJgo1gJASgADoRU7s7pXche
# MBK9Rp6103a50g5rmQzSM7TNsQIDAQABo4IBXTCCAVkwEgYDVR0TAQH/BAgwBgEB
# /wIBADAdBgNVHQ4EFgQUuhbZbU2FL3MpdpovdYxqII+eyG8wHwYDVR0jBBgwFoAU
# 7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoG
# CCsGAQUFBwMIMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYYaHR0cDovL29j
# c3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0cy5kaWdp
# Y2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNydDBDBgNVHR8EPDA6MDig
# NqA0hjJodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9v
# dEc0LmNybDAgBgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZI
# hvcNAQELBQADggIBAH1ZjsCTtm+YqUQiAX5m1tghQuGwGC4QTRPPMFPOvxj7x1Bd
# 4ksp+3CKDaopafxpwc8dB+k+YMjYC+VcW9dth/qEICU0MWfNthKWb8RQTGIdDAiC
# qBa9qVbPFXONASIlzpVpP0d3+3J0FNf/q0+KLHqrhc1DX+1gtqpPkWaeLJ7giqzl
# /Yy8ZCaHbJK9nXzQcAp876i8dU+6WvepELJd6f8oVInw1YpxdmXazPByoyP6wCeC
# RK6ZJxurJB4mwbfeKuv2nrF5mYGjVoarCkXJ38SNoOeY+/umnXKvxMfBwWpx2cYT
# gAnEtp/Nh4cku0+jSbl3ZpHxcpzpSwJSpzd+k1OsOx0ISQ+UzTl63f8lY5knLD0/
# a6fxZsNBzU+2QJshIUDQtxMkzdwdeDrknq3lNHGS1yZr5Dhzq6YBT70/O3itTK37
# xJV77QpfMzmHQXh6OOmc4d0j/R0o08f56PGYX/sr2H7yRp11LB4nLCbbbxV7HhmL
# NriT1ObyF5lZynDwN7+YAN8gFk8n+2BnFqFmut1VwDophrCYoCvtlUG3OtUVmDG0
# YgkPCr2B2RP+v6TR81fZvAT6gt4y3wSJ8ADNXcL50CN/AAvkdgIm2fBldkKmKYcJ
# RyvmfxqkhQ/8mJb2VVQrH4D6wPIOK+XW+6kvRBVK5xMOHds3OBqhK/bt1nz8MIIG
# uTCCBKGgAwIBAgIRAJmjgAomVTtlq9xuhKaz6jkwDQYJKoZIhvcNAQEMBQAwgYAx
# CzAJBgNVBAYTAlBMMSIwIAYDVQQKExlVbml6ZXRvIFRlY2hub2xvZ2llcyBTLkEu
# MScwJQYDVQQLEx5DZXJ0dW0gQ2VydGlmaWNhdGlvbiBBdXRob3JpdHkxJDAiBgNV
# BAMTG0NlcnR1bSBUcnVzdGVkIE5ldHdvcmsgQ0EgMjAeFw0yMTA1MTkwNTMyMTha
# Fw0zNjA1MTgwNTMyMThaMFYxCzAJBgNVBAYTAlBMMSEwHwYDVQQKExhBc3NlY28g
# RGF0YSBTeXN0ZW1zIFMuQS4xJDAiBgNVBAMTG0NlcnR1bSBDb2RlIFNpZ25pbmcg
# MjAyMSBDQTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAJ0jzwQwIzvB
# RiznM3M+Y116dbq+XE26vest+L7k5n5TeJkgH4Cyk74IL9uP61olRsxsU/WBAElT
# MNQI/HsE0uCJ3VPLO1UufnY0qDHG7yCnJOvoSNbIbMpT+Cci75scCx7UsKK1fcJo
# 4TXetu4du2vEXa09Tx/bndCBfp47zJNsamzUyD7J1rcNxOw5g6FJg0ImIv7nCeNn
# 3B6gZG28WAwe0mDqLrvU49chyKIc7gvCjan3GH+2eP4mYJASflBTQ3HOs6JGdriS
# MVoD1lzBJobtYDF4L/GhlLEXWgrVQ9m0pW37KuwYqpY42grp/kSYE4BUQrbLgBMN
# KRvfhQPskDfZ/5GbTCyvlqPN+0OEDmYGKlVkOMenDO/xtMrMINRJS5SY+jWCi8PR
# HAVxO0xdx8m2bWL4/ZQ1dp0/JhUpHEpABMc3eKax8GI1F03mSJVV6o/nmmKqDE6T
# K34eTAgDiBuZJzeEPyR7rq30yOVw2DvetlmWssewAhX+cnSaaBKMEj9O2GgYkPJ1
# 6Q5Da1APYO6n/6wpCm1qUOW6Ln1J6tVImDyAB5Xs3+JriasaiJ7P5KpXeiVV/HIs
# W3ej85A6cGaOEpQA2gotiUqZSkoQUjQ9+hPxDVb/Lqz0tMjp6RuLSKARsVQgETwo
# NQZ8jCeKwSQHDkpwFndfCceZ/OfCUqjxAgMBAAGjggFVMIIBUTAPBgNVHRMBAf8E
# BTADAQH/MB0GA1UdDgQWBBTddF1MANt7n6B0yrFu9zzAMsBwzTAfBgNVHSMEGDAW
# gBS2oVQ5AsOgP46KvPrU+Bym0ToO/TAOBgNVHQ8BAf8EBAMCAQYwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwMAYDVR0fBCkwJzAloCOgIYYfaHR0cDovL2NybC5jZXJ0dW0u
# cGwvY3RuY2EyLmNybDBsBggrBgEFBQcBAQRgMF4wKAYIKwYBBQUHMAGGHGh0dHA6
# Ly9zdWJjYS5vY3NwLWNlcnR1bS5jb20wMgYIKwYBBQUHMAKGJmh0dHA6Ly9yZXBv
# c2l0b3J5LmNlcnR1bS5wbC9jdG5jYTIuY2VyMDkGA1UdIAQyMDAwLgYEVR0gADAm
# MCQGCCsGAQUFBwIBFhhodHRwOi8vd3d3LmNlcnR1bS5wbC9DUFMwDQYJKoZIhvcN
# AQEMBQADggIBAHWIWA/lj1AomlOfEOxD/PQ7bcmahmJ9l0Q4SZC+j/v09CD2csX8
# Yl7pmJQETIMEcy0VErSZePdC/eAvSxhd7488x/Cat4ke+AUZZDtfCd8yHZgikGuS
# 8mePCHyAiU2VSXgoQ1MrkMuqxg8S1FALDtHqnizYS1bIMOv8znyJjZQESp9RT+6N
# H024/IqTRsRwSLrYkbFq4VjNn/KV3Xd8dpmyQiirZdrONoPSlCRxCIi54vQcqKiF
# LpeBm5S0IoDtLoIe21kSw5tAnWPazS6sgN2oXvFpcVVpMcq0C4x/CLSNe0XckmmG
# sl9z4UUguAJtf+5gE8GVsEg/ge3jHGTYaZ/MyfujE8hOmKBAUkVa7NMxRSB1EdPF
# pNIpEn/pSHuSL+kWN/2xQBJaDFPr1AX0qLgkXmcEi6PFnaw5T17UdIInA58rTu3m
# efNuzUtse4AgYmxEmJDodf8NbVcU6VdjWtz0e58WFZT7tST6EWQmx/OoHPelE77l
# ojq7lpsjhDCzhhp4kfsfszxf9g2hoCtltXhCX6NqsqwTT7xe8LgMkH4hVy8L1h2p
# qGLT2aNCx7h/F95/QvsTeGGjY7dssMzq/rSshFQKLZ8lPb8hFTmiGDJNyHga5hZ5
# 9IGynk08mHhBFM/0MLeBzlAQq1utNjQprztZ5vv/NJy8ua9AGbwkMWkOMIIGwjCC
# BKqgAwIBAgIQBUSv85SdCDmmv9s/X+VhFjANBgkqhkiG9w0BAQsFADBjMQswCQYD
# VQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lD
# ZXJ0IFRydXN0ZWQgRzQgUlNBNDA5NiBTSEEyNTYgVGltZVN0YW1waW5nIENBMB4X
# DTIzMDcxNDAwMDAwMFoXDTM0MTAxMzIzNTk1OVowSDELMAkGA1UEBhMCVVMxFzAV
# BgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMSAwHgYDVQQDExdEaWdpQ2VydCBUaW1lc3Rh
# bXAgMjAyMzCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAKNTRYcdg45b
# rD5UsyPgz5/X5dLnXaEOCdwvSKOXejsqnGfcYhVYwamTEafNqrJq3RApih5iY2nT
# WJw1cb86l+uUUI8cIOrHmjsvlmbjaedp/lvD1isgHMGXlLSlUIHyz8sHpjBoyoNC
# 2vx/CSSUpIIa2mq62DvKXd4ZGIX7ReoNYWyd/nFexAaaPPDFLnkPG2ZS48jWPl/a
# Q9OE9dDH9kgtXkV1lnX+3RChG4PBuOZSlbVH13gpOWvgeFmX40QrStWVzu8IF+qC
# ZE3/I+PKhu60pCFkcOvV5aDaY7Mu6QXuqvYk9R28mxyyt1/f8O52fTGZZUdVnUok
# L6wrl76f5P17cz4y7lI0+9S769SgLDSb495uZBkHNwGRDxy1Uc2qTGaDiGhiu7xB
# G3gZbeTZD+BYQfvYsSzhUa+0rRUGFOpiCBPTaR58ZE2dD9/O0V6MqqtQFcmzyrzX
# xDtoRKOlO0L9c33u3Qr/eTQQfqZcClhMAD6FaXXHg2TWdc2PEnZWpST618RrIbro
# HzSYLzrqawGw9/sqhux7UjipmAmhcbJsca8+uG+W1eEQE/5hRwqM/vC2x9XH3mwk
# 8L9CgsqgcT2ckpMEtGlwJw1Pt7U20clfCKRwo+wK8REuZODLIivK8SgTIUlRfgZm
# 0zu++uuRONhRB8qUt+JQofM604qDy0B7AgMBAAGjggGLMIIBhzAOBgNVHQ8BAf8E
# BAMCB4AwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAgBgNV
# HSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEwHwYDVR0jBBgwFoAUuhbZbU2F
# L3MpdpovdYxqII+eyG8wHQYDVR0OBBYEFKW27xPn783QZKHVVqllMaPe1eNJMFoG
# A1UdHwRTMFEwT6BNoEuGSWh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2Vy
# dFRydXN0ZWRHNFJTQTQwOTZTSEEyNTZUaW1lU3RhbXBpbmdDQS5jcmwwgZAGCCsG
# AQUFBwEBBIGDMIGAMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5j
# b20wWAYIKwYBBQUHMAKGTGh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydFRydXN0ZWRHNFJTQTQwOTZTSEEyNTZUaW1lU3RhbXBpbmdDQS5jcnQwDQYJ
# KoZIhvcNAQELBQADggIBAIEa1t6gqbWYF7xwjU+KPGic2CX/yyzkzepdIpLsjCIC
# qbjPgKjZ5+PF7SaCinEvGN1Ott5s1+FgnCvt7T1IjrhrunxdvcJhN2hJd6PrkKoS
# 1yeF844ektrCQDifXcigLiV4JZ0qBXqEKZi2V3mP2yZWK7Dzp703DNiYdk9WuVLC
# tp04qYHnbUFcjGnRuSvExnvPnPp44pMadqJpddNQ5EQSviANnqlE0PjlSXcIWiHF
# tM+YlRpUurm8wWkZus8W8oM3NG6wQSbd3lqXTzON1I13fXVFoaVYJmoDRd7ZULVQ
# jK9WvUzF4UbFKNOt50MAcN7MmJ4ZiQPq1JE3701S88lgIcRWR+3aEUuMMsOI5lji
# tts++V+wQtaP4xeR0arAVeOGv6wnLEHQmjNKqDbUuXKWfpd5OEhfysLcPTLfddY2
# Z1qJ+Panx+VPNTwAvb6cKmx5AdzaROY63jg7B145WPR8czFVoIARyxQMfq68/qTr
# eWWqaNYiyjvrmoI1VygWy2nyMpqy0tg6uLFGhmu6F/3Ed2wVbK6rr3M66ElGt9V/
# zLY4wNjsHPW2obhDLN9OTH0eaHDAdwrUAuBcYLso/zjlUlrWrBciI0707NMX+1Br
# /wd3H3GXREHJuEbTbDJ8WC9nR2XlG3O2mflrLAZG70Ee8PBf4NvZrZCARK+AEEGK
# MIIG3TCCBMWgAwIBAgIQJJHKH5ST1/a4Ov2c99Kv6zANBgkqhkiG9w0BAQsFADBW
# MQswCQYDVQQGEwJQTDEhMB8GA1UEChMYQXNzZWNvIERhdGEgU3lzdGVtcyBTLkEu
# MSQwIgYDVQQDExtDZXJ0dW0gQ29kZSBTaWduaW5nIDIwMjEgQ0EwHhcNMjQwOTAy
# MTA1NzI2WhcNMjUwOTAyMTA1NzI1WjCBgjELMAkGA1UEBhMCREUxEDAOBgNVBAgM
# B0JhdmFyaWExDzANBgNVBAcMBk11bmljaDEeMBwGA1UECgwVT3BlbiBTb3VyY2Ug
# RGV2ZWxvcGVyMTAwLgYDVQQDDCdPcGVuIFNvdXJjZSBEZXZlbG9wZXIsIEp1bGlh
# biBQYXdsb3dza2kwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCpFBRX
# Dt2Y6dui3ez8PJ3ql0CL0Whb4gs+3leX+q9nG1SWVabkIGvrPlEUQdzfJvPZM4PF
# 6m339OJKuS06IQwruBvnSGx7mcEgm82kwa6b0GRrodnbUm8WssnGUytY60YncjCu
# SgqfKI610e43oJMa+SVZCtXPHbiwciehhm93OJXyDOcevJhT/SXhh3zsEIq+hd1Y
# U5+XkKFXWUKD7ucxLB0BxXkAm2cj9wzkuTBCCi8RqTxW+m5b7VEJXcbjMk3HDDn7
# 5gom05FvYNmDTa2/fFDgh3BDUbTwsjbrvEoDdOySgAR3TIbkEVQaT4V/vxjy0xgr
# 6zpW4nP8XCqfwyPj4gbWTnEUrWd350IZcZbCcxwSfQkEqze2CxQiO9H706FLa7qD
# jP2iZH0OfkwhXntTbTNrYkNo2YmfSBN1LUPWV6w4PGRy1iaBcr3wOa1Re8nK8pon
# 9SaFvOfSoHhrzBFE8WfpBuMLRqN0W2C14b8N6qPWcpY2FRpj2KKdxhlnGMBjPh/p
# KM4Ltqg5ZZwGP1f5g4hodfVueaOuJXyNr7YnyEqD5YLcYCO9ycKCw4fgZ/cfxGmZ
# 2zmkSgo4SHZjoroGEs3QxOZVk2WN4as3YDeMqlgKNUOl2K8GoVonyYze8+yf1AYb
# 99q3p5tcObm4USL/hPqjxub5B52ga6uCoV2IeQIDAQABo4IBeDCCAXQwDAYDVR0T
# AQH/BAIwADA9BgNVHR8ENjA0MDKgMKAuhixodHRwOi8vY2NzY2EyMDIxLmNybC5j
# ZXJ0dW0ucGwvY2NzY2EyMDIxLmNybDBzBggrBgEFBQcBAQRnMGUwLAYIKwYBBQUH
# MAGGIGh0dHA6Ly9jY3NjYTIwMjEub2NzcC1jZXJ0dW0uY29tMDUGCCsGAQUFBzAC
# hilodHRwOi8vcmVwb3NpdG9yeS5jZXJ0dW0ucGwvY2NzY2EyMDIxLmNlcjAfBgNV
# HSMEGDAWgBTddF1MANt7n6B0yrFu9zzAMsBwzTAdBgNVHQ4EFgQUJ0fj+UjpDK75
# NKerDl0427zkmskwSwYDVR0gBEQwQjAIBgZngQwBBAEwNgYLKoRoAYb2dwIFAQQw
# JzAlBggrBgEFBQcCARYZaHR0cHM6Ly93d3cuY2VydHVtLnBsL0NQUzATBgNVHSUE
# DDAKBggrBgEFBQcDAzAOBgNVHQ8BAf8EBAMCB4AwDQYJKoZIhvcNAQELBQADggIB
# AJHuX55vDsuy3Ct90pCbzMtCQvuhA7RPPl+UC8JGheL23Dpx/2bFSoNpEPJK/bWe
# 0ah2gTlmMIp+92Tw45JesBY7jLUatkK8jIea+CwmGZSEV6lqY9nzX32nbpH0TLtk
# H72M6ns2A8pszloRdJ+GwhkyqIBXFWlGu5wx0rQQ258JarkF93syl4OXwQKeYhfP
# hSh0WL46+C3Nh90QH55T41/yb2RMl1PNqSv4n2Ev+e27mNIjlPUvvVqRqAYaM6OU
# LaJI+YOLQRPGPv7Np/oqnZ7Md5rMEH+v9TGERgEqDSRtuvfY1Te69HZhr2kFKDMm
# z4NSBy9YmONsgIMHSxP/PKZVKyxH7stwrzabbLyRdvSi+oHRVdG8KGilZ5ztsgOQ
# HJO88CQTaajLzzrshtf0eTEbkuXAB9RELw8eT2b4FD9/QTpx+63yPIz1O9jlY/cP
# /+LrdaysvDBtZOrAF9Gc8bgwG803w8Me3R3eN577NspJMxnoNlg2ZCmKnA+oq+YS
# NoYFNUEK8tdyD/JOv1HTpJRS61nevrNZ8MkH1vMFRx7aItRJSipIoYjaHcENm2Iq
# 3s5Sjvh+Qx9PibzZK/1ixwikthixF88uF7YxRWYbMqQO936+gAejVy3J+WetojTc
# zLRPmpC4mLDEKNOdFRq2/2VTQcizgn1/CUkbU5yUg57mMYIGQDCCBjwCAQEwajBW
# MQswCQYDVQQGEwJQTDEhMB8GA1UEChMYQXNzZWNvIERhdGEgU3lzdGVtcyBTLkEu
# MSQwIgYDVQQDExtDZXJ0dW0gQ29kZSBTaWduaW5nIDIwMjEgQ0ECECSRyh+Uk9f2
# uDr9nPfSr+swDQYJYIZIAWUDBAIBBQCggYQwGAYKKwYBBAGCNwIBDDEKMAigAoAA
# oQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4w
# DAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgeqrLYIZRtWwXJF1g82tf+4Z2
# devARQJT+T8nTxWO3hswDQYJKoZIhvcNAQEBBQAEggIAPjIlPrkWiHS/8vU0fQMb
# yT9TKfZuhTz2tn9YOyl6wPDOXwH808hPRc0W+yOjVU+/lv+lAGH3m0+27f/uX7+A
# qiDGegqIUTGNOko0Np0BE0RsgaXJ3fxUn8Y6Ub8lZuPEYeyxu5uO7A1iL9q03fQv
# EQTw/UbzsIX+aSLnFmoG57UwdQKniRAvx0AWfX52HZy0GxYjzkVPa6tF2KROAupz
# 6xBgtuQli2VXgRJlHvOZLredtCiXXEznbsgWYCUoSt/u+EyEOqZDDLCqS++RZUuy
# YVzYac23/K17oyEmnyFbpKgWE5q9+RTwF82pYjYu86q7KJquMQ5QHvpQk3G7l6Id
# pftuTzPR+Qq9VlTUcvp/K54HLyzOz6ZFIuqSbyPVrIkiXnFN/QQc3SMbCc/Z1Y8H
# g3QWTVw9eGzuZmKWBRc+AzNCwOqRPXJTWb9IUgnmTXDs03SKvV7yIQ4JGy0MHPbK
# X54E/Lq4gjBSZ9WebyRCmv2WHyeZ9CA5UjLfA0l5G6POENIjYeYLijVcpSinr3c4
# yeGx+kbN6VtHgDnYSBhxQktQ78sUVBL0+5BtiJ/jBapGml5lvDuO0uOOXmga6d3c
# wmBUPl8NwZ580eiLZ6zXJvSu98zrapXGAL8m69cBy/d1nwztKXj/RVbr/KYvaV33
# Z8+ivGpfRoIE+AjSatnBrEGhggMgMIIDHAYJKoZIhvcNAQkGMYIDDTCCAwkCAQEw
# dzBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xOzA5BgNV
# BAMTMkRpZ2lDZXJ0IFRydXN0ZWQgRzQgUlNBNDA5NiBTSEEyNTYgVGltZVN0YW1w
# aW5nIENBAhAFRK/zlJ0IOaa/2z9f5WEWMA0GCWCGSAFlAwQCAQUAoGkwGAYJKoZI
# hvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMjQwOTAyMTUxMDM2
# WjAvBgkqhkiG9w0BCQQxIgQg0QEGjQoJpVVPOFm7Xe72tn9nzsR+QkfniwGAC9bn
# neIwDQYJKoZIhvcNAQEBBQAEggIAJgQbxHqPfsopfuRpT4AnMOw+e0AFTnm4Tm8K
# AYm4PAq8ycWUob8tKo8gjuctz7MgBxYNu+vTnQv4jfOjVwVLbRv0ZO0NpSR+nPB/
# zJEjYU+Ra/GuLreuewWueTaxtWs0Vb6LMSQIONxk1acwGW/SgbufJ2ybPETV+LJ4
# iuqFrfYbSqYXR5bAvl3MOoaw67bJS8EhinfDtzBy3R98ebSmp+n8NJ6bxpVz0ka/
# ANw0QXzKXoOi6hRN6RjlVRFPbtGwIsFW6Cyts6UGwF7FDchm0BDxNgN1LSVUzUrt
# VXraGCsykWQuw1An8ZQtFc6pD8vINdjVhMKKHASLkAzezLzIPnaaqDraa2nhhgb+
# nO76+sk4msu87dS0od3xHxsdt1Zj/bFImHJbKnJiRubp5933EJcJ129S+zd8W0Js
# kN/z6LAYmC+a2DLnUeNnhQQQKrSZBaKURJ3gVHT4OIRfufhvtztBLZYK4Q8gPnfG
# 35GD9Ocfp7v1m+NhoitftLX6l0Qu6mxKiEFUvpRZ1e9PIcUWZeGD7lN0oFYFPZUZ
# xOoowKT/3Wi/Dk3FIJr4Qt8zg/UUrTLJFFhLhLLOFSl8k2vGx7dY6204hwfFO0XV
# PZFAU0LnducZfOKPr6/lYdZCy117tSqaASjr6m2JDC8g4ob5q5UIx5E3ktBywdrf
# UY+Z2WM=
# SIG # End signature block
