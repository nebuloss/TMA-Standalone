param(
    [string]$OutputDirectory,
    [string]$PayloadDirectory = (Join-Path (Split-Path $PSScriptRoot -Parent) "payload")
)

$ErrorActionPreference = "Stop"
$repositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$root = Join-Path $repositoryRoot 'src\TmaCleanRoom'
$output = if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    Join-Path $root "bin"
} else {
    $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDirectory)
}
New-Item -ItemType Directory -Path $output -Force | Out-Null
$csc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$outlook = Get-ChildItem "$env:WINDIR\assembly\GAC_MSIL\Microsoft.Office.Interop.Outlook" -Recurse -Filter Microsoft.Office.Interop.Outlook.dll | Select-Object -First 1 -ExpandProperty FullName
$office = Get-ChildItem "$env:WINDIR\assembly\GAC_MSIL\office" -Recurse -Filter OFFICE.DLL | Select-Object -First 1 -ExpandProperty FullName
$extensibility = Get-ChildItem "$env:WINDIR\assembly\GAC\Extensibility" -Recurse -Filter extensibility.dll | Select-Object -First 1 -ExpandProperty FullName
$windowsWinmd = Get-ChildItem "${env:ProgramFiles(x86)}\Windows Kits\10\UnionMetadata" -Recurse -Filter Windows.winmd -ErrorAction SilentlyContinue |
    Where-Object { $_.DirectoryName -notlike '*\Facade' } | Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName
$windowsRuntime = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\System.Runtime.WindowsRuntime.dll"
$frameworkFacades = "${env:ProgramFiles(x86)}\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8.1\Facades"
$systemRuntimeFacade = Join-Path $frameworkFacades "System.Runtime.dll"
$windowsInteropFacade = Join-Path $frameworkFacades "System.Runtime.InteropServices.WindowsRuntime.dll"
if (-not (Test-Path $csc) -or -not $outlook -or -not $office -or -not $extensibility -or
    -not $windowsWinmd -or -not (Test-Path $windowsRuntime) -or
    -not (Test-Path $systemRuntimeFacade) -or -not (Test-Path $windowsInteropFacade)) {
    throw "Office x64 or Windows Runtime build prerequisites are missing."
}
& $csc /nologo /target:library /optimize+ `
    "/out:$output\TmaCleanRoom.Addin.dll" `
    "/reference:$outlook" "/reference:$office" "/reference:$extensibility" `
    "/reference:$windowsWinmd" "/reference:$windowsRuntime" `
    "/reference:$systemRuntimeFacade" "/reference:$windowsInteropFacade" `
    /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:System.Web.Extensions.dll `
    (Join-Path $root "TmaCleanRoom.Addin.cs") `
    (Join-Path $root "OfficeNativeSignIn.cs") `
    (Join-Path $root "CleanRoomMeetingService.cs") `
    (Join-Path $root "LegacyTeamsSchedulerBridge.cs")
if ($LASTEXITCODE -ne 0) { throw "Clean-room compilation failed: $LASTEXITCODE" }
$templateSource = Join-Path $root "Templates\MeetingInvite.html"
if (-not (Test-Path -LiteralPath $templateSource)) {
    throw "Meeting invitation template is missing: $templateSource"
}
$templateOutput = Join-Path $output "Templates"
New-Item -ItemType Directory -Path $templateOutput -Force | Out-Null
Copy-Item -LiteralPath $templateSource -Destination $templateOutput -Force
$englishTemplateSource = Join-Path $root "Templates\MeetingInvite.en-US.html"
if (-not (Test-Path -LiteralPath $englishTemplateSource)) {
    throw "English meeting invitation template is missing: $englishTemplateSource"
}
Copy-Item -LiteralPath $englishTemplateSource -Destination $templateOutput -Force
$assetSource = Join-Path $PayloadDirectory "Assets\NewMeeting_Large_96.png"
if (-not (Test-Path -LiteralPath $assetSource)) {
    throw "Stock Teams ribbon asset is missing: $assetSource"
}
$assetOutput = Join-Path $output "Assets"
New-Item -ItemType Directory -Path $assetOutput -Force | Out-Null
Copy-Item -LiteralPath $assetSource -Destination $assetOutput -Force
Add-Type -AssemblyName System.Drawing
$meetNowPath = Join-Path $assetOutput "MeetNow_Large_96.png"
$bitmap = New-Object Drawing.Bitmap 32,32
$graphics = [Drawing.Graphics]::FromImage($bitmap)
try {
    $graphics.SmoothingMode = [Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.Clear([Drawing.Color]::Transparent)
    $brush = New-Object Drawing.SolidBrush ([Drawing.Color]::FromArgb(91,95,199))
    try {
        $graphics.FillRectangle($brush, 3, 7, 19, 18)
        $points = [Drawing.Point[]]@(
            [Drawing.Point]::new(22,12), [Drawing.Point]::new(30,8),
            [Drawing.Point]::new(30,24), [Drawing.Point]::new(22,20))
        $graphics.FillPolygon($brush, $points)
    } finally { $brush.Dispose() }
    $bitmap.Save($meetNowPath, [Drawing.Imaging.ImageFormat]::Png)
} finally { $graphics.Dispose(); $bitmap.Dispose() }
Get-Item "$output\TmaCleanRoom.Addin.dll" | Select-Object FullName,Length,LastWriteTime
