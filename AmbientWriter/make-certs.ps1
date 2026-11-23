# Create a local CA + a server certificate for this machine's hostname / LAN IP.
# ASCII only - may run under Windows PowerShell 5.1.
param(
  [string]$Dir = "$env:LOCALAPPDATA\AmbientWriter\certs",
  [string]$HostName = $env:COMPUTERNAME,
  [string]$Ip = "10.42.0.37",
  [string]$PfxPassword = "ambient"
)

$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $Dir | Out-Null

$sha256 = [System.Security.Cryptography.HashAlgorithmName]::SHA256
$pkcs1 = [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
$CertReq = [System.Security.Cryptography.X509Certificates.CertificateRequest]
$Store = [System.Security.Cryptography.X509Certificates.X509ContentType]

# ---- root CA ----
$caKey = [System.Security.Cryptography.RSA]::Create(2048)
$caReq = New-Object $CertReq("CN=AmbientWriter Local CA", $caKey, $sha256, $pkcs1)
$caReq.CertificateExtensions.Add((New-Object System.Security.Cryptography.X509Certificates.X509BasicConstraintsExtension($true, $false, 0, $true)))
$caReq.CertificateExtensions.Add((New-Object System.Security.Cryptography.X509Certificates.X509KeyUsageExtension(
  ([System.Security.Cryptography.X509Certificates.X509KeyUsageFlags]::KeyCertSign -bor
   [System.Security.Cryptography.X509Certificates.X509KeyUsageFlags]::CrlSign -bor
   [System.Security.Cryptography.X509Certificates.X509KeyUsageFlags]::DigitalSignature), $true)))
$caCert = $caReq.CreateSelfSigned([DateTimeOffset]::Now.AddDays(-1), [DateTimeOffset]::Now.AddYears(5))

[IO.File]::WriteAllBytes((Join-Path $Dir 'ca.cer'), $caCert.Export($Store::Cert))
"ca thumb = $($caCert.Thumbprint)"

# ---- server cert ----
$leafKey = [System.Security.Cryptography.RSA]::Create(2048)
$leafReq = New-Object $CertReq("CN=$HostName", $leafKey, $sha256, $pkcs1)
$leafReq.CertificateExtensions.Add((New-Object System.Security.Cryptography.X509Certificates.X509BasicConstraintsExtension($false, $false, 0, $true)))

$sanb = New-Object System.Security.Cryptography.X509Certificates.SubjectAlternativeNameBuilder
$sanb.AddDnsName($HostName)
$sanb.AddDnsName("$HostName.local")
if ($Ip) { $sanb.AddIpAddress([System.Net.IPAddress]::Parse($Ip)) }
$leafReq.CertificateExtensions.Add($sanb.Build())

$oids = New-Object System.Security.Cryptography.OidCollection
$oids.Add((New-Object System.Security.Cryptography.Oid("1.3.6.1.5.5.7.3.1"))) | Out-Null   # server auth
$leafReq.CertificateExtensions.Add((New-Object System.Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension($oids, $false)))

$serial = New-Object byte[] 8
(New-Object System.Random).NextBytes($serial)
$signed = $leafReq.Create($caCert, [DateTimeOffset]::Now.AddDays(-1), [DateTimeOffset]::Now.AddYears(2), $serial)
$leaf = $signed.CopyWithPrivateKey($leafKey)

[IO.File]::WriteAllBytes((Join-Path $Dir 'server.pfx'), $leaf.Export($Store::Pkcs12, $PfxPassword))
"leaf thumb = $($leaf.Thumbprint)"
"files:"
Get-ChildItem $Dir | Select-Object Name, Length | Format-Table -AutoSize | Out-String
