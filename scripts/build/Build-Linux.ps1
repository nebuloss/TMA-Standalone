param(
    [string]$OutputMsi,
    [switch]$AcceptNewMicrosoftVersion
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
if ($IsWindows) { throw 'Ce script est réservé au build Linux.' }

$root = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
if (-not $OutputMsi) { $OutputMsi = Join-Path $root 'TMA-Standalone.msi' }
$work = Join-Path $root '_work-linux'
$msix = Join-Path $work 'MSTeams-x64.msix'
$sourceMsi = Join-Path $work 'MicrosoftTeamsMeetingAddinInstaller.msi'
$extract = Join-Path $work 'extracted'
$payload = Join-Path $work 'payload'
$addin = Join-Path $work 'addin'
$filesWxs = Join-Path $work 'TmaFiles.wxs'
$lock = Get-Content (Join-Path $root 'dependencies.lock.json') -Raw | ConvertFrom-Json

foreach ($command in 'curl','unzip','msiextract','dotnet','python3','msiinfo') {
    if (-not (Get-Command $command -ErrorAction SilentlyContinue)) { throw "Prérequis Linux absent : $command" }
}
Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Path $work,$extract,$payload,$addin -Force | Out-Null

Write-Host '==> Téléchargement du MSIX x64 officiel Microsoft' -ForegroundColor Cyan
& curl -fL --retry 3 'https://go.microsoft.com/fwlink/?linkid=2196106' -o $msix
if ($LASTEXITCODE) { throw 'Téléchargement MSIX impossible.' }
& unzip -j $msix 'MicrosoftTeamsMeetingAddinInstaller.msi' -d $work
if ($LASTEXITCODE -or -not (Test-Path $sourceMsi)) { throw 'Le MSI TMA est absent du MSIX Microsoft.' }
$hash = (Get-FileHash $sourceMsi -Algorithm SHA256).Hash
if (-not $AcceptNewMicrosoftVersion -and $hash -ne $lock.microsoftTeamsMeetingAddinInstaller.sha256) {
    throw "MSI Microsoft différent du verrou : $hash"
}

Write-Host '==> Extraction du MSI avec msitools' -ForegroundColor Cyan
& msiextract -C $extract $sourceMsi
if ($LASTEXITCODE) { throw 'Extraction MSI impossible.' }
$loader = Get-ChildItem $extract -Recurse -Filter Microsoft.Teams.AddinLoader.dll |
    Where-Object FullName -Match '[/\\]x64[/\\]' | Select-Object -First 1
if (-not $loader) { throw 'Payload x64 introuvable.' }
$sourcePayload = $loader.Directory.FullName
$required = @(
 'Microsoft.Teams.MeetingAddin.dll','Microsoft.Teams.MeetingAddin.dll.config',
 'Microsoft.Teams.MeetingAddin.resources.dll','Microsoft.Teams.AuthLib.dll',
 'Microsoft.Teams.Diagnostics.dll','Microsoft.Applications.Telemetry.Windows.dll',
 'Microsoft.IdentityModel.JsonWebTokens.dll','Microsoft.IdentityModel.Logging.dll',
 'Microsoft.IdentityModel.Tokens.dll','Newtonsoft.Json.dll','OneAuth.dll',
 'System.IdentityModel.Tokens.Jwt.dll','System.Net.Http.Formatting.dll','adal.dll',
 'msvcp140.dll','vcruntime140.dll','vcruntime140_1.dll')
foreach ($name in $required) {
    $source = Join-Path $sourcePayload $name
    if (-not (Test-Path $source)) { throw "Dépendance Microsoft absente : $name" }
    Copy-Item $source $payload
}
New-Item -ItemType Directory (Join-Path $payload 'Assets') | Out-Null
Copy-Item (Join-Path $sourcePayload 'Assets/NewMeeting_Large_96.png') (Join-Path $payload 'Assets')

Write-Host '==> Compilation .NET Framework reproductible' -ForegroundColor Cyan
& dotnet build (Join-Path $root 'src/TmaCleanRoom/TmaCleanRoom.Addin.csproj') -c Release `
    -o $addin --nologo -p:ContinuousIntegrationBuild=true
if ($LASTEXITCODE) { throw 'Compilation du complément impossible.' }
Copy-Item (Join-Path $addin 'TmaCleanRoom.Addin.dll') $payload
New-Item -ItemType Directory (Join-Path $payload 'Templates') | Out-Null
Copy-Item (Join-Path $root 'src/TmaCleanRoom/Templates/*.html') (Join-Path $payload 'Templates')
Copy-Item (Join-Path $payload 'Assets/NewMeeting_Large_96.png') `
    (Join-Path $payload 'Assets/MeetNow_Large_96.png')

Write-Host '==> Génération WiX et compilation MSI' -ForegroundColor Cyan
$env:TMA_PAYLOAD = $payload
$env:TMA_WXS = $filesWxs
& python3 (Join-Path $root 'tools/generate-wix-files.py')
if ($LASTEXITCODE) { throw 'Génération WiX impossible.' }
& wix build (Join-Path $root 'installer/Product.wxs') $filesWxs -arch x64 -o $OutputMsi
if ($LASTEXITCODE -or -not (Test-Path $OutputMsi)) { throw 'Compilation MSI impossible.' }
$registry = (& msiinfo export $OutputMsi Registry | Out-String)
foreach ($requiredIdentity in 'TmaCleanRoom.Connect','8F5373B8-4973-4E58-A69E-CB57AA22691C') {
    if (-not $registry.Contains($requiredIdentity)) { throw "Identité MSI absente : $requiredIdentity" }
}
foreach ($forbidden in 'TeamsAddin.FastConnect','19A6E644-14E6-4A60-B8D7-DD20610A871D') {
    if ($registry.Contains($forbidden)) { throw "Identité stock interdite dans le MSI : $forbidden" }
}
Write-Host "MSI Linux valide : $OutputMsi" -ForegroundColor Green
