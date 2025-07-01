
<#
.SYNOPSIS
    Creates a self-signed certificate for Exchange Online App Authentication (Certificate-Based Authentication).

.DESCRIPTION
    This script generates a new self-signed certificate in the local certificate store (CurrentUser\My).
    It exports:
    - A public certificate (.cer) for uploading to Azure App Registration.
    - A private key (.pfx) for backup or future use.

    You will be prompted to set a password for the .pfx export.
    It also displays the generated certificate Thumbprint.

.NOTES
    - The certificate validity is set to 3 years.
    - The private key is exportable.
    - This script is intended for use with Microsoft Graph and Exchange Online Automation (Unattended Scripts).
    - Run PowerShell as the same user intended for certificate use.

.AUTHOR
    Mohammed Omar
#>

# ===============================
# Configuration
# ===============================

# Subject Name for the Certificate
$CertSubjectName = "CN=ExchangeGraphSyncApp"

# Get the path of the current script
$BasePath = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Create 'Certificates' folder under script path (optional, organized)
$CertFolder = Join-Path $BasePath "Certificates"
if (!(Test-Path $CertFolder)) { New-Item -Path $CertFolder -ItemType Directory -Force }

# Define paths for certificate export
$PublicCertPath = Join-Path $CertFolder "ExchangeGraphSyncApp.cer"
$PrivateCertPath = Join-Path $CertFolder "ExchangeGraphSyncApp.pfx"

# ===============================
# Create New Self-Signed Certificate
# ===============================

$Cert = New-SelfSignedCertificate `
    -Subject $CertSubjectName `
    -CertStoreLocation "cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(3)

# ===============================
# Export Public Certificate (.CER)
# ===============================

Export-Certificate -Cert $Cert -FilePath $PublicCertPath

# ===============================
# Export Private Certificate (.PFX)
# ===============================

# Prompt user for password to protect the .pfx file
$Password = Read-Host -AsSecureString "Enter password for PFX export"

# Export the private key
Export-PfxCertificate -Cert $Cert -FilePath $PrivateCertPath -Password $Password

# ===============================
# Display Certificate Details
# ===============================

Write-Host ""
Write-Host "Certificate Thumbprint:" $Cert.Thumbprint -ForegroundColor Green
Write-Host "Public Certificate (.cer) saved to:" $PublicCertPath -ForegroundColor Green
Write-Host "Private Certificate (.pfx) saved to:" $PrivateCertPath -ForegroundColor Yellow
Write-Host ""
