param([Parameter(Mandatory)][string]$OutputDirectory,[Parameter(Mandatory)][string]$PayloadDirectory,[string]$DotNet='dotnet',[string]$OfficePiaDirectory)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$project = Join-Path $root 'src\TmaCleanRoom\TmaCleanRoom.Addin.csproj'
if (-not $OfficePiaDirectory) { $OfficePiaDirectory = Join-Path $root '.tools\office-pia\ref' }
$office = Join-Path $OfficePiaDirectory 'office.dll'
$extensibility = Join-Path $OfficePiaDirectory 'Extensibility.dll'
if (-not (Test-Path $office) -or -not (Test-Path $extensibility)) {
    throw "PIA Office requises absentes de $OfficePiaDirectory. Exécutez scripts/dependencies/Get-OfficePia.ps1."
}
$piaProperty = '-p:OfficePiaDir=' + $OfficePiaDirectory
$check = Join-Path ([IO.Path]::GetTempPath()) ('tma-repro-' + [Guid]::NewGuid().ToString('N'))
try {
    New-Item -ItemType Directory $OutputDirectory,$check -Force | Out-Null
    & $DotNet build $project -c Release -o $OutputDirectory --nologo --no-incremental -p:RestoreLockedMode=true $piaProperty
    if ($LASTEXITCODE) { throw 'Première compilation .NET impossible.' }
    & $DotNet build $project -c Release -o $check --nologo --no-restore --no-incremental $piaProperty
    if ($LASTEXITCODE) { throw 'Compilation de contrôle impossible.' }
    $first = (Get-FileHash (Join-Path $OutputDirectory 'TmaCleanRoom.Addin.dll') -Algorithm SHA256).Hash
    $second = (Get-FileHash (Join-Path $check 'TmaCleanRoom.Addin.dll') -Algorithm SHA256).Hash
    if ($first -ne $second) { throw "DLL non reproductible : $first != $second" }
    Write-Host "DLL reproductible SHA-256 : $first" -ForegroundColor Green
}
finally { Remove-Item $check -Recurse -Force -ErrorAction SilentlyContinue }
$templates = Join-Path $OutputDirectory 'Templates'
New-Item -ItemType Directory $templates -Force | Out-Null
Copy-Item (Join-Path $root 'src\TmaCleanRoom\Templates\*.html') $templates -Force
$assets = Join-Path $OutputDirectory 'Assets'
New-Item -ItemType Directory $assets -Force | Out-Null
$sourceAssets = Join-Path $root 'src\TmaCleanRoom\Assets'
foreach ($icon in 'NewMeeting_Large_96.png','MeetNow_Large_96.png') {
    $sourceIcon = Join-Path $sourceAssets $icon
    if (-not (Test-Path $sourceIcon)) { throw "Icone requise absente : $sourceIcon" }
    Copy-Item $sourceIcon $assets -Force
}
