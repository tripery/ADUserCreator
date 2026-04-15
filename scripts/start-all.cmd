@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0start.ps1" -UiMode Docker %*
exit /b %errorlevel%
