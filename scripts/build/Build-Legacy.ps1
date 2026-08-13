param(
    [Parameter(Mandatory = $false)]
    [string]$MsiPath,
    [string]$OutputMsi
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
if (-not $MsiPath) { $MsiPath = Join-Path $root 'vendor\MicrosoftTeamsMeetingAddinInstaller.msi' }
if (-not $OutputMsi) { $OutputMsi = Join-Path $root 'TMA-Standalone.msi' }
$work = Join-Path $root "_work"
$extract = Join-Path $work "extracted"
$payload = Join-Path $work "payload"
$cleanRoomOutput = Join-Path $work "clean-room"
$filesWxs = Join-Path $work "TmaFiles.wxs"

function Step([string]$message) {
    Write-Host "`n==> $message" -ForegroundColor Cyan
}

function StableGuid([string]$text) {
    $md5 = [Security.Cryptography.MD5]::Create()
    try { return [Guid]::new($md5.ComputeHash([Text.Encoding]::UTF8.GetBytes($text.ToLowerInvariant()))) }
    finally { $md5.Dispose() }
}

function WixId([string]$prefix, [string]$value) {
    $safe = $value -replace '[^A-Za-z0-9_.]', '_'
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $suffix = (($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($value.ToLowerInvariant())))[0..5] | ForEach-Object ToString x2) -join '' }
    finally { $sha.Dispose() }
    $limit = 72 - $prefix.Length - $suffix.Length - 2
    if ($safe.Length -gt $limit) { $safe = $safe.Substring(0, $limit) }
    return "${prefix}_${safe}_${suffix}"
}

Step "Validation des prérequis"
if (-not (Test-Path -LiteralPath $MsiPath)) {
    throw "MSI Microsoft absent : $MsiPath. Lancez .\scripts\dependencies\Get-TmaInstaller.ps1 ou passez -MsiPath."
}
$signature = Get-AuthenticodeSignature -LiteralPath $MsiPath
if ($signature.Status -ne 'Valid' -or
    $signature.SignerCertificate.Subject -notmatch 'Microsoft Corporation') {
    throw "Le MSI d'entrée n'a pas une signature Authenticode Microsoft valide."
}
$lock = Get-Content -LiteralPath (Join-Path $root 'dependencies.lock.json') -Raw | ConvertFrom-Json
$actualHash = (Get-FileHash -LiteralPath $MsiPath -Algorithm SHA256).Hash
if ($actualHash -ne $lock.microsoftTeamsMeetingAddinInstaller.sha256) {
    Write-Warning "Version Microsoft non verrouillée : build valide mais non reproductible à l'identique."
}
foreach ($path in @($MsiPath, (Join-Path $root "installer\Product.wxs"),
    (Join-Path $root "scripts\build\Build-CleanRoom.ps1"))) {
    if (-not (Test-Path -LiteralPath $path)) { throw "Fichier requis introuvable : $path" }
}
if (-not (Get-Command wix.exe -ErrorAction SilentlyContinue)) { throw "WiX 4 (wix.exe) est absent du PATH." }

Step "Préparation du répertoire de travail"
Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Path $extract,$payload,$cleanRoomOutput -Force | Out-Null

Step "Extraction administrative du MSI Microsoft"
$log = Join-Path $work "msi-extract.log"
$arguments = @('/a', ('"' + (Resolve-Path $MsiPath).Path + '"'),
    ('TARGETDIR="' + $extract + '"'), '/qn', '/L*V', ('"' + $log + '"'))
$process = Start-Process msiexec.exe -ArgumentList $arguments -Wait -PassThru
if ($process.ExitCode -notin 0,3010) { throw "Échec extraction MSI ($($process.ExitCode)). Journal : $log" }

Step "Sélection du payload x64"
$loader = Get-ChildItem $extract -Recurse -Filter Microsoft.Teams.AddinLoader.dll |
    Where-Object FullName -Match '[\\/]x64[\\/]' | Select-Object -First 1
if (-not $loader) { throw "Payload Teams x64 introuvable dans le MSI source." }
$sourcePayload = $loader.Directory.FullName
$requiredMicrosoftFiles = @(
    'Microsoft.Teams.MeetingAddin.dll',
    'Microsoft.Teams.MeetingAddin.dll.config',
    'Microsoft.Teams.MeetingAddin.resources.dll',
    'Microsoft.Teams.AuthLib.dll',
    'Microsoft.Teams.Diagnostics.dll',
    'Microsoft.Applications.Telemetry.Windows.dll',
    'Microsoft.IdentityModel.JsonWebTokens.dll',
    'Microsoft.IdentityModel.Logging.dll',
    'Microsoft.IdentityModel.Tokens.dll',
    'Newtonsoft.Json.dll',
    'OneAuth.dll',
    'System.IdentityModel.Tokens.Jwt.dll',
    'System.Net.Http.Formatting.dll',
    'adal.dll',
    'msvcp140.dll',
    'vcruntime140.dll',
    'vcruntime140_1.dll'
)
foreach ($name in $requiredMicrosoftFiles) {
    $source = Join-Path $sourcePayload $name
    if (-not (Test-Path $source)) { throw "Composant Microsoft requis absent : $name" }
    Copy-Item -LiteralPath $source -Destination $payload
}
$payloadAssets = Join-Path $payload 'Assets'
New-Item -ItemType Directory -Path $payloadAssets -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $sourcePayload 'Assets\NewMeeting_Large_96.png') `
    -Destination $payloadAssets
foreach ($name in 'Microsoft.Teams.MeetingAddin.dll','Microsoft.Teams.AuthLib.dll','OneAuth.dll') {
    if (-not (Test-Path (Join-Path $payload $name))) { throw "Composant Microsoft requis absent : $name" }
}

Step "Compilation du complément COM indépendant"
& (Join-Path $root "scripts\build\Build-CleanRoom.ps1") `
    -OutputDirectory $cleanRoomOutput -PayloadDirectory $payload
if (-not (Test-Path (Join-Path $cleanRoomOutput "TmaCleanRoom.Addin.dll"))) { throw "DLL clean-room non produite." }
Copy-Item (Join-Path $cleanRoomOutput '*') $payload -Recurse -Force

Step "Génération déterministe de la liste WiX"
$files = Get-ChildItem $payload -Recurse -File | Sort-Object FullName
$directories = @{}
foreach ($file in $files) {
    $relative = $file.FullName.Substring($payload.Length).TrimStart('\')
    $directory = Split-Path $relative -Parent
    $current = ''
    foreach ($part in @($directory -split '\\' | Where-Object { $_ })) {
        $current = if ($current) { "$current\$part" } else { $part }
        if (-not $directories.ContainsKey($current)) {
            $directories[$current] = [pscustomobject]@{
                Path=$current; Parent=(Split-Path $current -Parent); Name=$part; Id=(WixId 'Dir' $current)
            }
        }
    }
}
$builder = [Text.StringBuilder]::new()
[void]$builder.AppendLine('<?xml version="1.0" encoding="utf-8"?>')
[void]$builder.AppendLine('<Wix xmlns="http://wixtoolset.org/schemas/v4/wxs"><Fragment><DirectoryRef Id="INSTALLFOLDER">')
function WriteDirectories([string]$parent,[int]$indent) {
    foreach ($directory in $directories.Values | Where-Object Parent -EQ $parent | Sort-Object Name) {
        [void]$builder.AppendLine((' ' * $indent) + '<Directory Id="' + $directory.Id + '" Name="' + [Security.SecurityElement]::Escape($directory.Name) + '">')
        WriteDirectories $directory.Path ($indent + 2)
        [void]$builder.AppendLine((' ' * $indent) + '</Directory>')
    }
}
WriteDirectories '' 4
[void]$builder.AppendLine('</DirectoryRef></Fragment><Fragment><ComponentGroup Id="TmaFiles">')
foreach ($file in $files) {
    $relative = $file.FullName.Substring($payload.Length).TrimStart('\')
    $relativeDirectory = Split-Path $relative -Parent
    $directoryId = if ($relativeDirectory) { $directories[$relativeDirectory].Id } else { 'INSTALLFOLDER' }
    $componentId = WixId 'Cmp' $relative
    $fileId = WixId 'File' $relative
    $source = [Security.SecurityElement]::Escape($file.FullName)
    [void]$builder.AppendLine("<Component Id=`"$componentId`" Guid=`"$(StableGuid $relative)`" Directory=`"$directoryId`"><File Id=`"$fileId`" Source=`"$source`" KeyPath=`"yes`" /></Component>")
}
[void]$builder.AppendLine('</ComponentGroup></Fragment></Wix>')
[IO.File]::WriteAllText($filesWxs, $builder.ToString(), [Text.UTF8Encoding]::new($false))

Step "Compilation du MSI"
Remove-Item -LiteralPath $OutputMsi -Force -ErrorAction SilentlyContinue
& wix.exe build (Join-Path $root "installer\Product.wxs") $filesWxs -arch x64 -o $OutputMsi
if ($LASTEXITCODE -ne 0 -or -not (Test-Path $OutputMsi)) { throw "Échec de compilation WiX." }

Step "Validation du MSI"
$installer = New-Object -ComObject WindowsInstaller.Installer
$database = $installer.GetType().InvokeMember('OpenDatabase','InvokeMethod',$null,$installer,@((Resolve-Path $OutputMsi).Path,0))
function ReadMsiColumn([string]$query) {
    $view=$database.GetType().InvokeMember('OpenView','InvokeMethod',$null,$database,@($query));
    $view.GetType().InvokeMember('Execute','InvokeMethod',$null,$view,$null); $values=@()
    while($record=$view.GetType().InvokeMember('Fetch','InvokeMethod',$null,$view,$null)){$values += $record.GetType().InvokeMember('StringData','GetProperty',$null,$record,@(1))}
    $view.GetType().InvokeMember('Close','InvokeMethod',$null,$view,$null); return $values
}
$registryView=$database.GetType().InvokeMember('OpenView','InvokeMethod',$null,$database,
    @('SELECT `Key`, `Name`, `Value` FROM `Registry`'))
$registryView.GetType().InvokeMember('Execute','InvokeMethod',$null,$registryView,$null)
$registry=@()
while($record=$registryView.GetType().InvokeMember('Fetch','InvokeMethod',$null,$registryView,$null)) {
    $registry += ((1..3 | ForEach-Object {
        $record.GetType().InvokeMember('StringData','GetProperty',$null,$record,@([int]$_))
    }) -join '|')
}
$registryView.GetType().InvokeMember('Close','InvokeMethod',$null,$registryView,$null)
foreach ($required in 'TmaCleanRoom.Addin','TmaCleanRoom.Connect','8F5373B8-4973-4E58-A69E-CB57AA22691C') {
    if (-not (($registry -join "`n").Contains($required))) { throw "Identité clean-room absente du MSI : $required" }
}
foreach ($forbidden in '19A6E644-14E6-4A60-B8D7-DD20610A871D','TeamsAddin.FastConnect') {
    if (($registry -join "`n").Contains($forbidden)) { throw "Le MSI tente de posséder une identité Teams stock : $forbidden" }
}
if ((Get-Item $OutputMsi).Length -lt 1MB) { throw "MSI anormalement petit." }
Write-Host "`nMSI valide : $OutputMsi" -ForegroundColor Green
