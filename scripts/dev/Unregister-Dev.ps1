$ErrorActionPreference = "Stop"
Remove-Item 'HKCU:\Software\Microsoft\Office\Outlook\Addins\TmaCleanRoom.Connect' -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item 'HKCU:\Software\Classes\TmaCleanRoom.Connect' -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item 'HKCU:\Software\Classes\CLSID\{8F5373B8-4973-4E58-A69E-CB57AA22691C}' -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "TMA Clean Room development registration removed."
