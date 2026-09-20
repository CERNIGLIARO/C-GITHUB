@echo off
setlocal
cd /d "%~dp0"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0SYNC_GOOGLE_VERS_GITHUB.ps1"

if errorlevel 1 (
  echo.
  echo La synchronisation CBC a rencontre une erreur.
  pause
  exit /b 1
)

echo.
echo Synchronisation CBC terminee.
exit /b 0
