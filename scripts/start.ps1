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

function Get-PreferredLanAddress {
    $candidate = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Where-Object {
            $_.IPAddress -notmatch '^127\.' -and
            $_.IPAddress -notmatch '^169\.254\.' -and
            $_.PrefixOrigin -ne 'WellKnown'
        } |
        Sort-Object -Property InterfaceMetric, SkipAsSource |
        Select-Object -First 1

    return $candidate.IPAddress
}

function Show-UiAccessInfo {
    param(
        [string]$UiPort = '5173',
        [string]$ApiPort = '8787'
    )

    $localUiUrl = "http://localhost:${UiPort}/"
    $localApiUrl = "http://localhost:${ApiPort}/api/health"
    $lanAddress = Get-PreferredLanAddress

    Write-Host ""
    Write-Host "Access URLs:" -ForegroundColor Cyan
    Write-Host "  UI (local): $localUiUrl" -ForegroundColor Green
    Write-Host "  API (local): $localApiUrl" -ForegroundColor Green
    if ($lanAddress) {
        Write-Host "  UI (LAN): http://${lanAddress}:${UiPort}/" -ForegroundColor Green
        Write-Host "  API (LAN): http://${lanAddress}:${ApiPort}/api/health" -ForegroundColor Green
    }
    Write-Host "  Ignore Docker/Vite container address like 172.x.x.x - that is internal only." -ForegroundColor Yellow
    Write-Host ""
}

function Test-DockerReady {
    try {
        docker info | Out-Null
        return $true
    } catch {
        return $false
    }
}

function Ensure-DockerRunning {
    if (Test-DockerReady) {
        Write-Host "Docker is ready." -ForegroundColor Green
        return
    }

    $dockerDesktopPath = Join-Path ${env:ProgramFiles} 'Docker\Docker\Docker Desktop.exe'
    if (-not (Test-Path $dockerDesktopPath)) {
        throw "Docker Desktop not found at '$dockerDesktopPath'. Start Docker manually or install Docker Desktop."
    }

    Write-Host "Starting Docker Desktop..." -ForegroundColor Cyan
    Start-Process -FilePath $dockerDesktopPath

    for ($attempt = 1; $attempt -le 30; $attempt++) {
        Start-Sleep -Seconds 2
        if (Test-DockerReady) {
            Write-Host "Docker is ready." -ForegroundColor Green
            return
        }
        Write-Host "Waiting for Docker Desktop (attempt $attempt/30)..." -ForegroundColor Yellow
    }

    throw "Docker Desktop did not become ready in time."
}

function Start-DockerUi {
    Write-Host "Starting Docker UI..." -ForegroundColor Cyan
    Ensure-DockerRunning
    Show-UiAccessInfo
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
    'Local' {
        Start-LocalUi
        Show-UiAccessInfo
    }
    'Docker' { Start-DockerUi }
}
