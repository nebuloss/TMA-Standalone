param([string]$Destination)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
if (-not $Destination) { $Destination = Join-Path $root '.tools\dotnet' }
$archive = Join-Path ([IO.Path]::GetTempPath()) 'dotnet-sdk-8.0.424-win-x64.zip'
$url = 'https://builds.dotnet.microsoft.com/dotnet/Sdk/8.0.424/dotnet-sdk-8.0.424-win-x64.zip'
$sha512 = '1787AB90635C2950672ED7C6507B000E1B212EA7D9A22FCEF37061344D37C64D4C4EDA12B8742601EFF5B45C8736485B31C55613892F240C300190E4E88A58B0'
try {
    Invoke-WebRequest $url -OutFile $archive
    if ((Get-FileHash $archive -Algorithm SHA512).Hash -ne $sha512) { throw 'Empreinte du SDK .NET invalide.' }
    New-Item -ItemType Directory $Destination -Force | Out-Null
    Expand-Archive $archive -DestinationPath $Destination -Force
}
finally { Remove-Item $archive -Force -ErrorAction SilentlyContinue }
Write-Host "SDK .NET reproductible prêt : $Destination" -ForegroundColor Green
