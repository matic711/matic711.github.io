@echo off
setlocal
net session >nul 2>&1
if %errorlevel% neq 0 (
  echo Zahtevam admin pravice...
  powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
  exit /b
)
powershell.exe -NoExit -NoProfile -ExecutionPolicy Bypass -File "%~dp0Setup-NewWorkstation.ps1" %*
