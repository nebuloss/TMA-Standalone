$ErrorActionPreference = "Stop"
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$dll = (Resolve-Path (Join-Path $root '_work\clean-room\TmaCleanRoom.Addin.dll')).Path
$clsid = "{8F5373B8-4973-4E58-A69E-CB57AA22691C}"
$class = "TmaCleanRoom.Addin"
$assembly = [Reflection.AssemblyName]::GetAssemblyName($dll).FullName
$inproc = "HKCU:\Software\Classes\CLSID\$clsid\InprocServer32"
New-Item -Path $inproc -Force | Out-Null
New-ItemProperty -Path $inproc -Name '(default)' -Value 'mscoree.dll' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $inproc -Name ThreadingModel -Value Both -PropertyType String -Force | Out-Null
New-ItemProperty -Path $inproc -Name Class -Value $class -PropertyType String -Force | Out-Null
New-ItemProperty -Path $inproc -Name Assembly -Value $assembly -PropertyType String -Force | Out-Null
New-ItemProperty -Path $inproc -Name RuntimeVersion -Value 'v4.0.30319' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $inproc -Name CodeBase -Value ([Uri]$dll).AbsoluteUri -PropertyType String -Force | Out-Null
$progid = 'HKCU:\Software\Classes\TmaCleanRoom.Connect'
New-Item -Path "$progid\CLSID" -Force | Out-Null
Set-Item -Path $progid -Value 'TMA Clean Room Outlook Add-in'
Set-Item -Path "$progid\CLSID" -Value $clsid
$addin = 'HKCU:\Software\Microsoft\Office\Outlook\Addins\TmaCleanRoom.Connect'
New-Item -Path $addin -Force | Out-Null
New-ItemProperty -Path $addin -Name FriendlyName -Value 'TMA Clean Room' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addin -Name Description -Value 'Add-in Outlook Teams autonome en C#' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addin -Name LoadBehavior -Value 3 -PropertyType DWord -Force | Out-Null
Write-Host "TMA Clean Room registered for the current user."
