$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function ConvertFrom-ReleaseTag([string]$Tag) {
    if ($Tag -notmatch '^v?(\d+)\.(\d+)\.(\d+)$') { return $null }
    return [version]::new(
        [int]$Matches[1], [int]$Matches[2], [int]$Matches[3])
}

# An explicit value is useful for reproducibility tests and downstream packagers.
if ($env:TMA_VERSION) {
    $explicit = ConvertFrom-ReleaseTag $env:TMA_VERSION
    if (-not $explicit) { throw "Version MSI invalide : $env:TMA_VERSION" }
    $explicit.ToString(3)
    exit 0
}

$root = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$tagsAtHead = @(git -C $root tag --points-at HEAD 2>$null)
foreach ($tag in $tagsAtHead) {
    $release = ConvertFrom-ReleaseTag $tag
    if ($release) {
        $release.ToString(3)
        exit 0
    }
}

$releaseTags = @(git -C $root tag --list 'v[0-9]*' --sort=-version:refname 2>$null)
foreach ($tag in $releaseTags) {
    $previous = ConvertFrom-ReleaseTag $tag
    if ($previous) {
        # Untagged builds represent the next patch candidate. Repeated CI builds
        # remain installable because the MSI permits same-version upgrades.
        ([version]::new($previous.Major, $previous.Minor,
            $previous.Build + 1)).ToString(3)
        exit 0
    }
}

# Bootstrap value for a new fork. Tags become the source of truth afterwards.
'0.1.0'
