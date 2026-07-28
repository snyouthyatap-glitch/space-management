@echo off
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0install-dashboard.ps1"
if errorlevel 1 (
  echo.
  echo Installation failed.
  pause
  exit /b 1
)
