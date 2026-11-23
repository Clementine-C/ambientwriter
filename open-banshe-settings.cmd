@echo off
rem Open the companion-writing settings page in the default browser.
rem Works without the Word add-in: it is the same panel served by the local HTTPS server.
setlocal
set REPO=%~dp0word-gpt-plus

powershell -NoProfile -Command "if (-not (Get-NetTCPConnection -LocalPort 3000 -State Listen -ErrorAction SilentlyContinue)) { Start-Process -WindowStyle Hidden -FilePath 'node' -ArgumentList 'serve-https.cjs' -WorkingDirectory '%REPO%'; Start-Sleep -Seconds 3 }"

start "" "https://%COMPUTERNAME%:3000/settings?tab=ambient"
