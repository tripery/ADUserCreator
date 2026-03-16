param(
    [ValidateSet('Local', 'Docker')]
    [string]$UiMode = 'Docker'
)

$scriptsRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptsRoot
$apiUrl = 'http://localhost:8787/api/health'

function Start-ApiServer {
    Write-Host "Starting PowerShell API..." -ForegroundColor Cyan
    $apiCommand = "cd `"$projectRoot\webapi`"; .\server.ps1"
    Start-Process powershell -ArgumentList "-NoExit", "-ExecutionPolicy", "Bypass", "-Command", $apiCommand
}

function Wait-ForApiHealth {
    Write-Host "Waiting for API health check at $apiUrl ..." -ForegroundColor DarkCyan
    for ($attempt = 1; $attempt -le 15; $attempt++) {
        Start-Sleep -Seconds 2
        try {
            $response = Invoke-WebRequest -UseBasicParsing $apiUrl -TimeoutSec 3
            if ($response.StatusCode -eq 200) {
                Write-Host "API is ready." -ForegroundColor Green
                return $true
            }
        } catch {
            Write-Host "API not ready yet (attempt $attempt/15)." -ForegroundColor Yellow
        }
    }

    Write-Warning "PowerShell API did not respond on $apiUrl. UI will still start, but /api requests will fail until webapi/server.ps1 is healthy."
    return $false
}

function Start-LocalUi {
    Write-Host "Starting local React UI..." -ForegroundColor Cyan
    $uiCommand = "cd `"$projectRoot\webui-react`"; npm run dev"
    Start-Process powershell -ArgumentList "-NoExit", "-ExecutionPolicy", "Bypass", "-Command", $uiCommand
}

function Start-DockerUi {
    Write-Host "Starting Docker UI..." -ForegroundColor Cyan
    Push-Location $projectRoot
    try {
        docker compose up --build
    } finally {
        Pop-Location
    }
}

Start-ApiServer
Wait-ForApiHealth | Out-Null

switch ($UiMode) {
    'Local' { Start-LocalUi }
    'Docker' { Start-DockerUi }
}
