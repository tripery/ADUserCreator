Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Resolve-AppRoot {
    $candidates = @()

    if ($PSScriptRoot) {
        $candidates += $PSScriptRoot
    }

    $baseDir = [System.AppDomain]::CurrentDomain.BaseDirectory.TrimEnd('\\')
    if ($baseDir) {
        $candidates += $baseDir
        $candidates += (Join-Path $baseDir 'src')
    }

    $seen = @{}
    foreach ($candidate in $candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        if ($seen.ContainsKey($candidate)) { continue }
        $seen[$candidate] = $true

        if (Test-Path (Join-Path $candidate 'ui\\Splash.ps1')) {
            return $candidate
        }
    }

    throw "Не знайдено файли застосунку (ui\\Splash.ps1). Перевірені шляхи: $($seen.Keys -join ', ')"
}

$root = Resolve-AppRoot

function Assert-Requirements {
    $requiredModules = @(
        'ActiveDirectory',
        'ImportExcel'
    )

    $requiredCommands = @(
        'Get-ADDomain',
        'Get-ADObject',
        'Get-ADUser',
        'Get-ADGroup',
        'Import-Excel',
        'Open-ExcelPackage',
        'Close-ExcelPackage'
    )

    $missingModules = @()
    foreach ($moduleName in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $moduleName)) {
            $missingModules += $moduleName
        }
    }

    # Try importing modules so exported commands become visible even in ISE.
    foreach ($moduleName in ($requiredModules | Where-Object { $_ -notin $missingModules })) {
        try {
            Import-Module $moduleName -ErrorAction Stop | Out-Null
        }
        catch {
            if ($moduleName -notin $missingModules) {
                $missingModules += $moduleName
            }
        }
    }

    $missingCommands = @()
    foreach ($commandName in $requiredCommands) {
        if (-not (Get-Command $commandName -ErrorAction SilentlyContinue)) {
            $missingCommands += $commandName
        }
    }

    if ($missingModules.Count -eq 0 -and $missingCommands.Count -eq 0) {
        return
    }

    $lines = @(
        "Не знайдено всі потрібні модулі/команди для запуску."
    )

    if ($missingModules.Count -gt 0) {
        $lines += ""
        $lines += "Відсутні модулі:"
        $lines += ($missingModules | ForEach-Object { " - $_" })
    }

    if ($missingCommands.Count -gt 0) {
        $lines += ""
        $lines += "Відсутні команди:"
        $lines += ($missingCommands | ForEach-Object { " - $_" })
    }

    $isISE = ($Host.Name -like '*ISE*') -or ($null -ne $psISE)
    $lines += ""
    if ($isISE) {
        $lines += "PowerShell ISE: запустіть x64 ISE від імені адміністратора і встановіть модулі:"
    } else {
        $lines += "Встановіть потрібні модулі у PowerShell x64 (краще від імені адміністратора):"
    }
    $lines += "Install-Module ActiveDirectory   # RSAT / AD module (через Windows Features)"
    $lines += "Install-Module ImportExcel -Scope CurrentUser"
    $lines += "Після встановлення перезапустіть ISE/EXE."

    throw ($lines -join [Environment]::NewLine)
}

try {
    Assert-Requirements

    # Splash
    . "$root\ui\Splash.ps1"
    $splash = Show-Splash

    # Core
    . "$root\common\Logging.ps1"
    . "$root\common\Password.ps1"

    # AD
    . "$root\ad\Transliteration.ps1"
    . "$root\ad\Naming.ps1"
    . "$root\ad\UserProvision.ps1"

    # Excel
    . "$root\excel\ExcelImport.ps1"

    # UI
    . "$root\ui\SelectOU.ps1"
    . "$root\ui\SelectGroups.ps1"
    . "$root\ui\Form.ps1"

    $splash.Close()

    Show-MainForm
}
catch {
    try { $splash.Close() } catch {}

    [System.Windows.Forms.MessageBox]::Show(
        $_.Exception.Message,
        "Помилка запуску",
        "OK",
        "Error"
    ) | Out-Null
}
