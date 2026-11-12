@echo off
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall-dashboard.ps1"
if errorlevel 1 (
  echo.
  echo Uninstall failed.
  pause
  exit /b 1
)
pause
