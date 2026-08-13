param(
    [string]$MsiPath,
    [string]$OutputMsi
)
$ErrorActionPreference = 'Stop'
if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
    throw 'Utilisez Build-Linux.ps1 sur Linux.'
}
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
if (-not $MsiPath) { $MsiPath = Join-Path $root 'vendor\MicrosoftTeamsMeetingAddinInstaller.msi' }
if (-not $OutputMsi) { $OutputMsi = Join-Path $root 'TMA-Standalone.msi' }
$env:TMA_VERSION = & (Join-Path $PSScriptRoot 'Get-PackageVersion.ps1')
$wixl = Join-Path $root '.tools\wixl'
$dotnet = Join-Path $root '.tools\dotnet\dotnet.exe'
if (-not (Test-Path $dotnet)) { & (Join-Path $root 'scripts\dependencies\Get-DotNetSdk.ps1') }
if (-not (Test-Path (Join-Path $wixl 'wixl.exe'))) {
    & (Join-Path $root 'scripts\dependencies\Get-WixlTools.ps1') -Destination $wixl
}
& (Join-Path $PSScriptRoot 'Build-Legacy.ps1') -MsiPath $MsiPath -OutputMsi $OutputMsi -WixlDirectory $wixl -DotNet $dotnet
if ($LASTEXITCODE) { exit $LASTEXITCODE }
