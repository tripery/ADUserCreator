$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "Starting PowerShell API..." -ForegroundColor Cyan
Start-Process powershell -ArgumentList "-NoExit", "-ExecutionPolicy", "Bypass", "-File", "`"$projectRoot\webapi\server.ps1`""

Start-Sleep -Seconds 2

Write-Host "Starting Docker UI..." -ForegroundColor Cyan
docker compose up --build
