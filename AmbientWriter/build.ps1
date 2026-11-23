# AmbientWriter: build + per-user COM registration (no admin required)
# NOTE: keep this file ASCII-only - it may be run by Windows PowerShell 5.1.

$ErrorActionPreference = 'Stop'

$root   = Split-Path -Parent $MyInvocation.MyCommand.Path
$src    = Join-Path $root 'src\AmbientWriter.cs'
$tpl    = Join-Path $root 'config.template.ini'
$outDir = Join-Path $env:LOCALAPPDATA 'AmbientWriter'
$dll    = Join-Path $outDir 'AmbientWriter.dll'
$csc    = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$pia    = 'C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Word.dll'
$office = 'C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\OFFICE.DLL'
$fx     = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319'

foreach ($p in @($csc, $pia, $office, $src, $tpl)) {
    if (-not (Test-Path $p)) { throw "missing dependency: $p" }
}

New-Item -ItemType Directory -Force -Path $outDir | Out-Null

Write-Host '== compile =='
& $csc /nologo /codepage:65001 /target:library /platform:anycpu /optimize+ /out:"$dll" `
    /r:"$pia" /r:"$office" `
    /r:"$fx\System.dll" /r:"$fx\System.Core.dll" /r:"$fx\System.Windows.Forms.dll" `
    /r:"$fx\System.Drawing.dll" /r:"$fx\System.Web.Extensions.dll" `
    "$src"
if ($LASTEXITCODE -ne 0) { throw "compile failed (exit $LASTEXITCODE)" }

Write-Host '== compile watcher (out-of-process) =='
$watcherSrc = Join-Path $root 'src\Watcher.cs'
$watcherExe = Join-Path $outDir 'AmbientWriterWatcher.exe'
& $csc /nologo /codepage:65001 /target:winexe /platform:anycpu /optimize+ /out:"$watcherExe" `
    /r:"$pia" /r:"$office" `
    /r:"$fx\System.dll" /r:"$fx\System.Core.dll" /r:"$fx\System.Windows.Forms.dll" `
    /r:"$fx\System.Drawing.dll" /r:"$fx\System.Web.Extensions.dll" `
    "$watcherSrc"
if ($LASTEXITCODE -ne 0) { throw "watcher compile failed (exit $LASTEXITCODE)" }

$asmName = [System.Reflection.AssemblyName]::GetAssemblyName($dll)

$clsid  = '{7C3E9A21-5B4D-4F0E-9A11-2D6C8B3F5E70}'
$progId = 'AmbientWriter.Connect'
$class  = 'AmbientWriter.Connect'

Write-Host '== write HKCU registry =='

$clsidKey = "HKCU:\Software\Classes\CLSID\$clsid"
New-Item -Path "$clsidKey\InprocServer32" -Force | Out-Null
Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name '(Default)' -Value 'mscoree.dll'
Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name 'ThreadingModel' -Value 'Both'
Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name 'Class' -Value $class
Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name 'Assembly' -Value $asmName.FullName
Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name 'RuntimeVersion' -Value 'v4.0.30319'
Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name 'CodeBase' -Value ("file:///" + $dll.Replace('\', '/'))
New-Item -Path "$clsidKey\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}" -Force | Out-Null

New-Item -Path "HKCU:\Software\Classes\$progId" -Force | Out-Null
Set-ItemProperty -Path "HKCU:\Software\Classes\$progId" -Name '(Default)' -Value 'AmbientWriter'
New-Item -Path "HKCU:\Software\Classes\$progId\CLSID" -Force | Out-Null
Set-ItemProperty -Path "HKCU:\Software\Classes\$progId\CLSID" -Name '(Default)' -Value $clsid

$addinKey = "HKCU:\Software\Microsoft\Office\Word\Addins\$progId"
New-Item -Path $addinKey -Force | Out-Null
Set-ItemProperty -Path $addinKey -Name 'FriendlyName' -Value 'AmbientWriter'
Set-ItemProperty -Path $addinKey -Name 'Description' -Value 'Inline grey ghost-text continuation for Word'
Set-ItemProperty -Path $addinKey -Name 'LoadBehavior' -Value 3 -Type DWord
Set-ItemProperty -Path $addinKey -Name 'CommandLineSafe' -Value 0 -Type DWord

$cfg = Join-Path $outDir 'config.ini'
if (-not (Test-Path $cfg)) { Copy-Item $tpl $cfg -Force }

Write-Host '== done ==' -ForegroundColor Green
Write-Host "  DLL    : $dll"
Write-Host "  config : $cfg"
Write-Host "  log    : $(Join-Path $outDir 'ambientwriter.log')"
Write-Host '  next   : open Word, type some text, pause ~1.1s'
