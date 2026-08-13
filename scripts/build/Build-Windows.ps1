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
& (Join-Path $PSScriptRoot 'Build-Legacy.ps1') -MsiPath $MsiPath -OutputMsi $OutputMsi
if ($LASTEXITCODE) { exit $LASTEXITCODE }
