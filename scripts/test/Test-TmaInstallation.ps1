$ErrorActionPreference = 'Stop'
$clsid = '{8F5373B8-4973-4E58-A69E-CB57AA22691C}'
$addinKey = 'HKCU:\Software\Microsoft\Office\Outlook\Addins\TmaCleanRoom.Connect'
$inprocKey = "HKCU:\Software\Classes\CLSID\$clsid\InprocServer32"
$results = [ordered]@{}
$results['Outlook x64'] = [Environment]::Is64BitOperatingSystem -and
    (Test-Path 'HKLM:\SOFTWARE\Microsoft\Office\Outlook')
$results['Add-in registration'] = Test-Path $addinKey
$results['LoadBehavior = 3'] = (Get-ItemProperty $addinKey -ErrorAction SilentlyContinue).LoadBehavior -eq 3
$results['Independent CLSID'] = Test-Path $inprocKey
$results['CLR in-process server'] = (Get-ItemProperty $inprocKey -ErrorAction SilentlyContinue).'(default)' -eq 'mscoree.dll'
$codeBase = (Get-ItemProperty $inprocKey -ErrorAction SilentlyContinue).CodeBase
$results['Installed add-in DLL'] = $codeBase -and (Test-Path ([Uri]$codeBase).LocalPath)
$results['Local Teams scheduler'] = Test-Path "$env:ProgramFiles\TMA-Standalone\Microsoft.Teams.MeetingAddin.dll"
$results['French invitation template'] = Test-Path "$env:ProgramFiles\TMA-Standalone\Templates\MeetingInvite.html"
$results['English invitation template'] = Test-Path "$env:ProgramFiles\TMA-Standalone\Templates\MeetingInvite.en-US.html"
$results['Meeting icon asset'] = Test-Path "$env:ProgramFiles\TMA-Standalone\Assets\NewMeeting_Large_96.png"
$results['Stock CLSID untouched'] = (Get-ItemProperty 'HKCU:\Software\Classes\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}\InprocServer32' -ErrorAction SilentlyContinue).'(default)' -notlike '*TMA-Standalone*'
$results.GetEnumerator() | ForEach-Object {
    [pscustomobject]@{ Check=$_.Key; Result=if ($_.Value) {'OK'} else {'FAIL'} }
} | Format-Table -AutoSize
if ($results.Values -contains $false) { exit 1 }
