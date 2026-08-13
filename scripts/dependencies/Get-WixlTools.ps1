param([string]$Destination)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
if (-not $Destination) { $Destination = Join-Path $root '.tools\wixl' }
$packages = @(
 @('msitools-0.106-4','F39EAFA75A09B7E47D87BF73DD02975B4AC115DC799E0B73C7B2471190839D8E'),
 @('glib2-2.88.1-1','ACDD349636D53C03739F584F1A23C96DE28F0F5FC26A24F4386970EFDBCA48AC'),
 @('gcab-1.6-3','32FACF670EC763562128C34B8AA2E75F4FF06FD95DC1C90ED652FF226FDB9D7C'),
 @('gettext-runtime-1.0-1','BA693DDA4AC375AF76CE481FF3A6E7481286546CC7DC6D56C7021DAE34084157'),
 @('libxml2-2.15.3-1','6E52E2D3F887098FF2FB98D3A4CA8FC8F1FD0AD0D6643F76B9A5D5C3E03019E1'),
 @('zlib-1.3.2-2','841401182976D2F9E17E5C0EBAAC51F2A8014140EA53D67625E91C8FB3C85EA0'),
 @('libiconv-1.19-1','9A500F38C2B91808741C62FAE746B3E9110B33A1ECF5C30FA0C66DBEDDDF7E16'),
 @('libffi-3.5.2-1','36DDB7F89C020E1A1A56B3633069A5F9ACE10B312C4A81A78198EE5D5D175C47'),
 @('libgsf-1.14.55-2','38BFDDC2885A5C040B8FE7DE4E8B2115E2E3F49612C8EB285172E27854C59878'),
 @('pcre2-10.47-1','839BC4642F94C44E94E331C9092C6D186B1EDC54DFDF6A81CB2062F638417023'),
 @('bzip2-1.0.8-3','932DA2C63B23E6A4448757EB36FB198A9E51121874408270AFE2C91BADA513C7')
)
$temporary = Join-Path ([IO.Path]::GetTempPath()) ('wixl-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory $temporary -Force | Out-Null
try {
 foreach ($package in $packages) {
  $file = "mingw-w64-ucrt-x86_64-$($package[0])-any.pkg.tar.zst"
  $archive = Join-Path $temporary $file
  Invoke-WebRequest "https://repo.msys2.org/mingw/ucrt64/$file" -OutFile $archive
  if ((Get-FileHash $archive -Algorithm SHA256).Hash -ne $package[1]) { throw "Empreinte invalide : $file" }
  tar.exe -xf $archive -C $temporary
  if ($LASTEXITCODE) { throw "Extraction impossible : $file" }
 }
 New-Item -ItemType Directory $Destination -Force | Out-Null
 $binaries = 'wixl.exe','msiinfo.exe','msiextract.exe','libglib-2.0-0.dll','libgio-2.0-0.dll',
  'libgcab-1.0-0.dll','libintl-8.dll','libgobject-2.0-0.dll','libxml2-16.dll',
  'libmsi-1.0-0.dll','zlib1.dll','libiconv-2.dll','libffi-8.dll',
  'libgsf-1-114.dll','libpcre2-8-0.dll','libgmodule-2.0-0.dll','libbz2-1.dll'
 foreach ($name in $binaries) { Copy-Item (Join-Path $temporary "ucrt64\bin\$name") $Destination -Force }
 Copy-Item (Join-Path $temporary 'ucrt64\share\wixl-0.106') (Join-Path $Destination 'share\wixl-0.106') -Recurse -Force
}
finally { Remove-Item $temporary -Recurse -Force -ErrorAction SilentlyContinue }
Write-Host "wixl portable prêt : $Destination" -ForegroundColor Green
