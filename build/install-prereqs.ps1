param(
    [switch]$Quiet
)

$ErrorActionPreference = 'Stop'

function Write-Info($m) { Write-Host "[INFO] $m" }
function Write-Ok($m)   { Write-Host "[OK]   $m" }
function Write-Warn($m) { Write-Host "[WARN] $m" }

function Ensure-ImportExcel {
    if (Get-Module -ListAvailable -Name ImportExcel) {
        try { Import-Module ImportExcel -ErrorAction Stop | Out-Null } catch {}
        Write-Ok 'ImportExcel already installed'
        return
    }

    Write-Info 'Installing ImportExcel from PowerShell Gallery...'

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    } catch {}

    try {
        $psGallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if ($psGallery -and $psGallery.InstallationPolicy -ne 'Trusted') {
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
        }
    } catch {
        Write-Warn "Failed to update PSGallery trust: $($_.Exception.Message)"
    }

    try {
        $nuget = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue
        if (-not $nuget) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
        }
    } catch {
        Write-Warn "NuGet provider setup warning: $($_.Exception.Message)"
    }

    Install-Module ImportExcel -Scope AllUsers -Force -AllowClobber -ErrorAction Stop
    Import-Module ImportExcel -ErrorAction Stop | Out-Null
    Write-Ok 'ImportExcel installed'
}

function Ensure-ActiveDirectoryModule {
    if (Get-Command Get-ADDomain -ErrorAction SilentlyContinue) {
        try { Import-Module ActiveDirectory -ErrorAction Stop | Out-Null } catch {}
        Write-Ok 'ActiveDirectory cmdlets already available'
        return
    }

    Write-Info 'ActiveDirectory cmdlets not found. Trying to install/enable RSAT AD tools...'

    $adInstalled = $false

    # Windows client (Windows 10/11): RSAT capability
    if (Get-Command Add-WindowsCapability -ErrorAction SilentlyContinue) {
        $capNames = @(
            'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0',
            'Rsat.ActiveDirectory*'
        )

        foreach ($cap in $capNames) {
            try {
                $caps = Get-WindowsCapability -Online -Name $cap -ErrorAction Stop
                foreach ($c in @($caps)) {
                    if ($c.State -eq 'Installed') {
                        Write-Ok "RSAT capability already installed: $($c.Name)"
                        $adInstalled = $true
                        break
                    }

                    Write-Info "Installing Windows capability: $($c.Name)"
                    Add-WindowsCapability -Online -Name $c.Name -ErrorAction Stop | Out-Null
                    $adInstalled = $true
                    break
                }
            } catch {
                Write-Warn "RSAT capability path failed for '$cap': $($_.Exception.Message)"
            }

            if ($adInstalled) { break }
        }
    }

    # Windows Server: feature install
    if (-not $adInstalled -and (Get-Command Install-WindowsFeature -ErrorAction SilentlyContinue)) {
        try {
            Write-Info 'Trying Install-WindowsFeature RSAT-AD-PowerShell...'
            Install-WindowsFeature RSAT-AD-PowerShell -IncludeAllSubFeature -ErrorAction Stop | Out-Null
            $adInstalled = $true
        } catch {
            Write-Warn "Install-WindowsFeature failed: $($_.Exception.Message)"
        }
    }

    Start-Sleep -Seconds 1

    try {
        Import-Module ActiveDirectory -ErrorAction Stop | Out-Null
    } catch {
        Write-Warn "Import-Module ActiveDirectory failed: $($_.Exception.Message)"
    }

    if (-not (Get-Command Get-ADDomain -ErrorAction SilentlyContinue)) {
        throw "Не вдалося автоматично підготувати ActiveDirectory cmdlets. Встановіть RSAT 'Active Directory Domain Services and Lightweight Directory Services Tools' та перезапустіть застосунок."
    }

    Write-Ok 'ActiveDirectory cmdlets available'
}

function Assert-FinalCommands {
    $required = @(
        'Get-ADDomain', 'Get-ADObject', 'Get-ADUser', 'Get-ADGroup',
        'Import-Excel', 'Open-ExcelPackage', 'Close-ExcelPackage'
    )

    $missing = @($required | Where-Object { -not (Get-Command $_ -ErrorAction SilentlyContinue) })
    if ($missing.Count -gt 0) {
        throw ("Після інсталяції все ще відсутні команди: " + ($missing -join ', '))
    }

    Write-Ok 'All required commands are available'
}

try {
    Write-Info 'Checking and installing prerequisites for ADUserCreator...'
    Ensure-ActiveDirectoryModule
    Ensure-ImportExcel
    Assert-FinalCommands
    Write-Ok 'Prerequisites check completed successfully'
    exit 0
}
catch {
    $msg = $_.Exception.Message
    Write-Host "[ERROR] $msg"

    if (-not $Quiet) {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        [System.Windows.Forms.MessageBox]::Show(
            $msg,
            'ADUserCreator prerequisites',
            'OK',
            'Error'
        ) | Out-Null
    }

    exit 1
}
