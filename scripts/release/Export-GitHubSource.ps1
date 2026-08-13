param([string]$Destination)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
if (-not $Destination) { $Destination = Join-Path $root 'github-export' }
if (Test-Path $Destination) { throw "Le repertoire de destination existe deja : $Destination" }
$allowed = @(
    '.gitignore','dependencies.lock.json','vendor\README.md','README.md','LICENSE','NOTICE.md',
    'SECURITY.md','CONTRIBUTING.md','installer\Product.wxs','tools\generate-wix-files.py','tools\generate-wixl.py',
    '.github\workflows\source-audit.yml','.github\workflows\linux-build.yml',
    'scripts\build\Build-Legacy.ps1','scripts\build\Build-CleanRoom.ps1',
    'scripts\build\Build-Windows.ps1','scripts\build\Build-Linux.ps1',
    'scripts\dependencies\Get-TmaInstaller.ps1','scripts\dev\Register-Dev.ps1',
    'scripts\dev\Unregister-Dev.ps1','scripts\test\Test-TmaInstallation.ps1',
    'scripts\test\Test-SourceRelease.ps1','scripts\release\Export-GitHubSource.ps1',
    'src\TmaCleanRoom\CleanRoomMeetingService.cs','src\TmaCleanRoom\LegacyTeamsSchedulerBridge.cs',
    'src\TmaCleanRoom\OfficeNativeSignIn.cs','src\TmaCleanRoom\TmaCleanRoom.Addin.cs',
    'src\TmaCleanRoom\TmaCleanRoom.Addin.csproj','src\TmaCleanRoom\ExtensibilityInterop.cs',
    'src\TmaCleanRoom\Templates\MeetingInvite.html','src\TmaCleanRoom\Templates\MeetingInvite.en-US.html'
)
New-Item -ItemType Directory -Path $Destination | Out-Null
foreach ($relative in $allowed) {
    $source = Join-Path $root $relative
    if (-not (Test-Path $source)) { throw "Source attendue absente : $relative" }
    $target = Join-Path $Destination $relative
    New-Item -ItemType Directory -Path (Split-Path $target -Parent) -Force | Out-Null
    Copy-Item -LiteralPath $source -Destination $target
}
if (Get-ChildItem $Destination -Recurse -File | Where-Object Extension -In '.dll','.exe','.msi') {
    throw 'Un binaire interdit est present dans export.'
}
git -C $Destination init | Out-Null
git -C $Destination add .
Write-Host "Export source GitHub cree : $Destination" -ForegroundColor Green
