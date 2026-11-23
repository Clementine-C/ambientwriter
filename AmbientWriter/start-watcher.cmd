@echo off
rem AmbientWriter watcher launcher (no console window)
set EXE=%LOCALAPPDATA%\AmbientWriter\AmbientWriterWatcher.exe
if not exist "%EXE%" (
  echo AmbientWriter watcher not found: %EXE%
  echo Run build.ps1 first.
  pause
  exit /b 1
)
start "" "%EXE%"
