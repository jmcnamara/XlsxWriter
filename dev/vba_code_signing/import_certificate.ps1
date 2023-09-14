# import certificate.pfx for current user
$pwd = ConvertTo-SecureString -String "xlsxwriter" -Force -AsPlainText
Import-PfxCertificate certificate.pfx -CertStoreLocation Cert:\CurrentUser\My -Exportable -Password $pwd

# import certificate.pfx also as trusted root certificate
Import-PfxCertificate certificate.pfx -CertStoreLocation Cert:\CurrentUser\Root -Exportable -Password $pwd
