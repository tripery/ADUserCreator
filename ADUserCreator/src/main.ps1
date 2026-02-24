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

try {

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
