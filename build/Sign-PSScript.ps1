param
(
    # Specifies the path to the script that is being signed.
    [Parameter(Mandatory = $true)]
    [string] $ScriptPath,
    # Specifies the certificate that will be used to sign the script or file.
    [Parameter(Mandatory = $false)]
    [X509Certificate] $SigningCertificate = (Get-ChildItem Cert:\CurrentUser\My\E7413D745138A6DC584530AECE27CEFDDA9D9CD6 -CodeSigningCert),
    # Uses the specified time stamp server to add a time stamp to the signature.
    [Parameter(Mandatory = $false)]
    [string] $TimestampServer = 'http://timestamp.digicert.com'
)

## Sign PowerShell Files
Set-AuthenticodeSignature $ScriptPath -Certificate $SigningCertificate -HashAlgorithm SHA256 -IncludeChain NotRoot -TimestampServer $TimestampServer
