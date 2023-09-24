# Generate new self-signed certificate for code signing (-Type CodeSigningCert) and add it to current user's certificates (Cert:\CurrentUser\My)
# The private key is allowed to be exported by specifying -KeyExportPolicy Exportable.
$cert = New-SelfSignedCertificate -CertStoreLocation Cert:\CurrentUser\My -Subject "CN=VBA Code Signing" -KeyAlgorithm RSA -KeyLength 2048 -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" -KeyExportPolicy Exportable -KeyUsage DigitalSignature -Type CodeSigningCert

# Export certificate and private key to .pfx file. The private key is protected by a password.
$pwd = ConvertTo-SecureString -String "xlsxwriter" -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath certificate.pfx -Password $pwd

# Import the certificate also as trusted root certificate.
# A popup window will occur when not running as administrator. The warning should be confirmed.
Import-PfxCertificate certificate.pfx -CertStoreLocation Cert:\CurrentUser\Root -Exportable -Password $pwd
