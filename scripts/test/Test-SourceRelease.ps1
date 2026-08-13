$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$wixlGenerator = Get-Content -LiteralPath (Join-Path $root 'tools\generate-wixl.py') -Raw
if ($wixlGenerator -notmatch '<MajorUpgrade Schedule="afterInstallInitialize"') {
    throw 'MajorUpgrade must remove the previous product before installing stable components.'
}
Push-Location $root
try {
$tracked = @(git ls-files)
$forbidden = @($tracked | Where-Object {
    $_ -like 'payload/*' -or $_ -match '\.(dll|exe|msi|pdb|wixpdb|cab)$'
})
if ($forbidden) {
    throw "Fichiers binaires interdits suivis par Git :`n$($forbidden -join "`n")"
}
$patterns = 'access_token','refresh_token','client_secret','BEGIN PRIVATE KEY'
foreach ($pattern in $patterns) {
    $hits = @(git grep -n -I -i $pattern -- ':!scripts/test/Test-SourceRelease.ps1' 2>$null)
    $grepExitCode = $LASTEXITCODE
    if ($grepExitCode -notin 0,1) { throw "git grep a échoué avec le code $grepExitCode." }
    if ($hits) { throw "Motif sensible trouvé ($pattern) :`n$($hits -join "`n")" }
}
$nativeIdentitySource = Get-Content 'src/TmaCleanRoom/OfficeNativeSignIn.cs' -Raw
foreach ($forbiddenNativeIdentityPattern in 'ValidatedOsfVersion',
    'OsfCurrentIdentityRva','ShowUiOrdinal','GetCurrentOfficeIdentity',
    'Marshal.WriteInt64(parameters','Marshal.WriteByte(parameters') {
    if ($nativeIdentitySource.Contains($forbiddenNativeIdentityPattern)) {
        throw "Dépendance Office versionnée interdite : $forbiddenNativeIdentityPattern"
    }
}
foreach ($required in 'README.md','LICENSE','NOTICE.md','SECURITY.md','dependencies.lock.json','vendor/README.md',
    'global.json','scripts/dependencies/Get-TmaInstaller.ps1','scripts/dependencies/Get-OfficePia.ps1',
    'scripts/dependencies/Get-DotNetSdk.ps1','scripts/build/Build-Managed.ps1','scripts/build/Build-Windows.ps1',
    'scripts/build/Build-Linux.ps1','scripts/build/Get-PackageVersion.ps1',
    'src/TmaCleanRoom/TmaCleanRoom.Addin.cs',
    'src/TmaCleanRoom/TmaCleanRoom.Addin.csproj') {
    if ($tracked -notcontains $required) { throw "Fichier de publication absent : $required" }
}

# Both localized templates are runtime inputs, not embedded resources. Validate
# their contract here so a design edit cannot produce a successful but unusable MSI.
$templateTokens = '{{JOIN_URL}}','{{MEETING_ID}}','{{TEAMS_ICON}}',
    '{{PASSCODE_ROW}}','{{OPTIONS_BLOCK}}'
foreach ($templatePath in 'src/TmaCleanRoom/Templates/MeetingInvite.html',
    'src/TmaCleanRoom/Templates/MeetingInvite.en-US.html') {
    if (-not (Test-Path -LiteralPath $templatePath)) {
        throw "Template d'invitation absent : $templatePath"
    }
    $template = Get-Content -LiteralPath $templatePath -Raw
    foreach ($token in $templateTokens) {
        if (-not $template.Contains($token)) {
            throw "Marqueur $token absent de $templatePath"
        }
    }
    if ($template -notmatch 'data-tma-clean-room="meeting"') {
        throw "Marqueur racine TMA absent de $templatePath"
    }
}
Write-Host 'Audit source-only réussi.' -ForegroundColor Green
$global:LASTEXITCODE = 0
}
finally { Pop-Location }
