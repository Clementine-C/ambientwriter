@echo off
rem One-shot repair for the GPT Plus add-in registration.
rem Run this, then add the add-in once in Word (Start tab -> Add-ins -> GPT Plus).
setlocal
set MANIFEST=%USERPROFILE%\Documents\WordGPT\manifest.xml
if not exist "%MANIFEST%" (
  echo Manifest missing: %MANIFEST%
  pause
  exit /b 1
)
echo Registering add-in...
call npx -y office-addin-dev-settings@latest register "%MANIFEST%"
echo.
echo Now in Word:
echo   Start tab -^> Add-ins -^> GPT Plus  (Developer add-ins section)
echo   then close Word and open it again.
echo.
echo Current registrations:
call npx -y office-addin-dev-settings@latest registered
pause