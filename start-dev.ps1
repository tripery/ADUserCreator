Start-Process powershell -ArgumentList "-NoExit -Command cd '$PSScriptRoot\webapi'; .\server.ps1"
Start-Process powershell -ArgumentList "-NoExit -Command cd '$PSScriptRoot\webui-react'; npm run dev"
