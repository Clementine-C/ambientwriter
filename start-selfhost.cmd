@echo off
rem Start the self-hosted Word GPT Plus instance over HTTPS (hostname cert).
rem Serves word-gpt-plus\dist on https://%COMPUTERNAME%:3000
set REPO=%~dp0word-gpt-plus
if not exist "%LOCALAPPDATA%\AmbientWriter\certs\server.pfx" (
  echo Missing certificate. Run: powershell -ExecutionPolicy Bypass -File "%~dp0AmbientWriter\make-certs.ps1"
  pause
  exit /b 1
)
start "AmbientWriter self-host" /min cmd /c "cd /d "%REPO%" && node serve-https.cjs"
