@echo off
rem Start the local HTTPS server + AmbientWriter, then open Word.
rem NOTE: GPT Plus itself needs to be added once by hand in Word
rem       (Start tab -> Add-ins -> GPT Plus); after that it loads per session.
setlocal
set REPO=%~dp0word-gpt-plus
set WATCHER=%LOCALAPPDATA%\AmbientWriter\AmbientWriterWatcher.exe

powershell -NoProfile -Command "if (-not (Get-NetTCPConnection -LocalPort 3000 -State Listen -ErrorAction SilentlyContinue)) { Start-Process -WindowStyle Hidden -FilePath 'node' -ArgumentList 'serve-https.cjs' -WorkingDirectory '%REPO%'; Start-Sleep -Seconds 3 }"

powershell -NoProfile -Command "if (-not (Get-Process AmbientWriterWatcher -ErrorAction SilentlyContinue)) { Start-Process -WindowStyle Hidden -FilePath '%WATCHER%'; Start-Sleep -Seconds 1 }"

start "" "%ProgramFiles%\Microsoft Office\root\Office16\WINWORD.EXE"

echo.
echo Started: local server + companion writing (grey inline suggestions).
echo GPT Plus: if it is not in the ribbon yet, add it once:
echo   Start tab -^> Add-ins -^> GPT Plus