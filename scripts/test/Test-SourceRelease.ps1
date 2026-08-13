$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
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
foreach ($required in 'README.md','LICENSE','NOTICE.md','SECURITY.md','dependencies.lock.json','vendor/README.md',
    'global.json','scripts/dependencies/Get-TmaInstaller.ps1','scripts/dependencies/Get-OfficePia.ps1',
    'scripts/dependencies/Get-DotNetSdk.ps1','scripts/build/Build-Managed.ps1','scripts/build/Build-Windows.ps1',
    'scripts/build/Build-Linux.ps1','src/TmaCleanRoom/TmaCleanRoom.Addin.cs',
    'src/TmaCleanRoom/TmaCleanRoom.Addin.csproj') {
    if ($tracked -notcontains $required) { throw "Fichier de publication absent : $required" }
}
Write-Host 'Audit source-only réussi.' -ForegroundColor Green
$global:LASTEXITCODE = 0
}
finally { Pop-Location }
