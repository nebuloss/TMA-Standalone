param(
    [string]$Destination,
    [switch]$ProvisionTeamsIfMissing,
    [switch]$AcceptNewVersion
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
if (-not $Destination) { $Destination = Join-Path $root 'vendor\MicrosoftTeamsMeetingAddinInstaller.msi' }

function Test-MicrosoftSignature([string]$Path) {
    $signature = Get-AuthenticodeSignature -LiteralPath $Path
    if ($signature.Status -ne 'Valid') {
        throw "Signature Authenticode invalide pour $Path : $($signature.Status)"
    }
    $subject = $signature.SignerCertificate.Subject
    if ($subject -notmatch 'Microsoft Corporation') {
        throw "Le signataire de $Path n'est pas Microsoft Corporation : $subject"
    }
}

function Find-TmaMsi {
    $candidates = New-Object Collections.Generic.List[string]
    $packages = @(Get-AppxPackage -Name MSTeams -ErrorAction SilentlyContinue)
    foreach ($package in $packages) {
        if ($package.InstallLocation) {
            $candidate = Join-Path $package.InstallLocation 'MicrosoftTeamsMeetingAddinInstaller.msi'
            $candidates.Add($candidate)
        }
    }
    $windowsApps = Join-Path $env:ProgramFiles 'WindowsApps'
    if (Test-Path $windowsApps) {
        Get-ChildItem $windowsApps -Directory -Filter 'MSTeams_*_x64__*' `
            -ErrorAction SilentlyContinue | Sort-Object Name -Descending |
            ForEach-Object {
                $candidate = Join-Path $_.FullName 'MicrosoftTeamsMeetingAddinInstaller.msi'
                $candidates.Add($candidate)
            }
    }
    return $candidates | Where-Object { Test-Path -LiteralPath $_ } |
        Select-Object -First 1
}

$source = Find-TmaMsi
if (-not $source -and $ProvisionTeamsIfMissing) {
    $temporaryName = 'TmaBootstrap-' + [Guid]::NewGuid().ToString('N')
    $temporary = Join-Path ([IO.Path]::GetTempPath()) $temporaryName
    New-Item -ItemType Directory -Path $temporary | Out-Null
    try {
        $bootstrapper = Join-Path $temporary 'teamsbootstrapper.exe'
        $uri = 'https://go.microsoft.com/fwlink/?clcid=0x409&linkid=2243204'
        Invoke-WebRequest -Uri $uri -OutFile $bootstrapper -UseBasicParsing
        Test-MicrosoftSignature $bootstrapper
        $process = Start-Process -FilePath $bootstrapper -ArgumentList '-p' `
            -Verb RunAs -Wait -PassThru -WindowStyle Hidden
        if ($process.ExitCode -ne 0) {
            throw ('TeamsBootstrapper a échoué avec HRESULT 0x{0:X8}.' -f
                ([uint32]$process.ExitCode))
        }
        Start-Sleep -Seconds 3
        $source = Find-TmaMsi
    }
    finally {
        Remove-Item -LiteralPath $temporary -Recurse -Force -ErrorAction SilentlyContinue
    }
}

if (-not $source) {
    throw @'
Le MSI TMA est introuvable dans le package New Teams.
Installez New Teams, ou relancez ce script avec -ProvisionTeamsIfMissing.
'@
}

Test-MicrosoftSignature $source
$lockPath = Join-Path $root 'dependencies.lock.json'
$lock = Get-Content -LiteralPath $lockPath -Raw | ConvertFrom-Json
$expectedHash = $lock.microsoftTeamsMeetingAddinInstaller.sha256
$sourceHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
if (-not $AcceptNewVersion -and $sourceHash -ne $expectedHash) {
    throw @"
La version Microsoft installée ne correspond pas à la release verrouillée.
Attendu : $expectedHash
Trouvé  : $sourceHash

Utilisez -AcceptNewVersion pour construire avec cette version. Ce choix rend
le build fonctionnel, mais pas byte-for-byte reproductible avec la release.
"@
}
$destinationDirectory = Split-Path $Destination -Parent
New-Item -ItemType Directory -Path $destinationDirectory -Force | Out-Null
Copy-Item -LiteralPath $source -Destination $Destination -Force
Test-MicrosoftSignature $Destination
$hash = Get-FileHash -LiteralPath $Destination -Algorithm SHA256
Write-Host "MSI Microsoft prêt : $Destination" -ForegroundColor Green
Write-Host "Source  : $source"
Write-Host "SHA-256 : $($hash.Hash)"
