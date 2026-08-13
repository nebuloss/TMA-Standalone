param([string]$Destination)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
if (-not $Destination) { $Destination = Join-Path $root '.tools\office-pia\ref' }
$expected = @{ office = 'DC1C9337435FA37201DBB8C012E0397E0A1BAE7273305CA397FEED566BA0F9E9'; extensibility = 'B342088C8A4403092703BF40062041265E12EDD204AFF4F6532226478A65CBB2' }
$office = Join-Path $Destination 'office.dll'; $extensibility = Join-Path $Destination 'Extensibility.dll'
if ((Test-Path $office) -and (Test-Path $extensibility) -and (Get-FileHash $office -Algorithm SHA256).Hash -eq $expected.office -and (Get-FileHash $extensibility -Algorithm SHA256).Hash -eq $expected.extensibility) { Write-Host "PIA Office déjà validées : $Destination" -ForegroundColor Green; return }
$temp = Join-Path ([IO.Path]::GetTempPath()) ('tma-pia-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory $temp,$Destination -Force | Out-Null
try {
 $nupkg=Join-Path $temp 'office.nupkg'; Invoke-WebRequest 'https://api.nuget.org/v3-flatcontainer/msofficecore.interop/15.0.2/msofficecore.interop.15.0.2.nupkg' -OutFile $nupkg
 if ((Get-FileHash $nupkg -Algorithm SHA256).Hash -ne 'FCE5ED291A692C5C498F66EFD5CB5A6E3970133A10087EB92152CE1023E41F5F') { throw 'Package office.dll non conforme.' }
 Expand-Archive $nupkg (Join-Path $temp 'nupkg') -Force; Copy-Item (Join-Path $temp 'nupkg\lib\net48\Office.dll') $office -Force
 $redist=Join-Path $temp 'PIARedist.exe'; Invoke-WebRequest 'https://download.microsoft.com/download/c/1/d/c1d6dbbb-700d-4669-98cf-820ac3ae8e55/PIARedist.exe' -OutFile $redist
 if ((Get-FileHash $redist -Algorithm SHA256).Hash -ne 'F1D04AAE514D9734231B05860DFBB644E8DA61C4E1F2A7F5B3E41D2FCA2FAEA9') { throw 'PIARedist Microsoft non conforme.' }
 if ($IsWindows) { $p=Start-Process $redist -ArgumentList ("/extract:$temp"),'/quiet' -Wait -PassThru; if($p.ExitCode){throw "PIARedist: $($p.ExitCode)"}; $msiInfo=Join-Path $root '.tools\wixl\msiinfo.exe' }
 else { foreach($c in '7z','msiinfo','cabextract'){if(-not(Get-Command $c -ErrorAction SilentlyContinue)){throw "Prérequis absent: $c"}}; & 7z x $redist ("-o$temp") -y | Out-Null; if($LASTEXITCODE){throw 'Extraction PIARedist impossible.'}; $msiInfo='msiinfo' }
 $msi=Get-ChildItem $temp -Recurse -Filter o2010pia.msi | Select-Object -First 1; if(-not $msi){throw 'o2010pia.msi absent.'}
 $cab=Join-Path $temp 'PIAREDIST.CAB'; $p=Start-Process $msiInfo -ArgumentList 'extract',$msi.FullName,'PIAREDIST.CAB' -RedirectStandardOutput $cab -Wait -PassThru; if($p.ExitCode){throw 'Cabinet PIA inaccessible.'}
 $name='FL_extensibility_dll_____X86.3643236F_FC70_11D3_A536_0090278A1BB8'
 if($IsWindows){& expand.exe $cab "-F:$name" $temp | Out-Null}else{& cabextract -q -F $name -d $temp $cab}; if($LASTEXITCODE){throw 'Extensibility.dll inaccessible.'}
 Copy-Item (Join-Path $temp $name) $extensibility -Force
 if((Get-FileHash $office -Algorithm SHA256).Hash -ne $expected.office -or (Get-FileHash $extensibility -Algorithm SHA256).Hash -ne $expected.extensibility){throw 'Empreinte finale des PIA invalide.'}
} finally { Remove-Item $temp -Recurse -Force -ErrorAction SilentlyContinue }
Write-Host "PIA Office de compilation validées : $Destination" -ForegroundColor Green
