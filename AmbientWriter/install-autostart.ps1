# AmbientWriter: one-time autostart registration for the current user (no admin).
$ErrorActionPreference = 'Stop'

$exe = Join-Path $env:LOCALAPPDATA 'AmbientWriter\AmbientWriterWatcher.exe'
if (-not (Test-Path $exe)) { throw "not built yet: $exe (run build.ps1 first)" }

$startup = [Environment]::GetFolderPath('Startup')
$lnk = Join-Path $startup 'AmbientWriter.lnk'
$ws = New-Object -ComObject WScript.Shell
$sc = $ws.CreateShortcut($lnk)
$sc.TargetPath = $exe
$sc.WorkingDirectory = Split-Path $exe
$sc.WindowStyle = 7
$sc.Description = 'AmbientWriter - inline AI continuation for Word'
$sc.Save()

Write-Host "autostart shortcut created: $lnk"
Write-Host "remove it (or delete the .lnk) to disable autostart."
