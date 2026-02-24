# build\build-exe.ps1
Set-Location $PSScriptRoot

if (-not (Get-Command ps2exe -ErrorAction SilentlyContinue)) {
    Install-Module ps2exe -Scope CurrentUser -Force
}

$projectRoot = (Resolve-Path '..').Path
$srcDirCandidate = Join-Path $projectRoot 'src'
$distDir = Join-Path $projectRoot 'dist'
$runtimeDirs = @('ad', 'common', 'excel', 'ui')
$runtimeFiles = @('main.ps1')

if (Test-Path (Join-Path $srcDirCandidate 'main.ps1')) {
    $srcDir = (Resolve-Path $srcDirCandidate).Path
} elseif (Test-Path (Join-Path $distDir 'main.ps1')) {
    $srcDir = (Resolve-Path $distDir).Path
} else {
    throw "Не знайдено main.ps1 ані в '$srcDirCandidate', ані в '$distDir'."
}

$srcMain = (Resolve-Path (Join-Path $srcDir 'main.ps1')).Path

if (-not (Test-Path $distDir)) {
    New-Item -ItemType Directory -Path $distDir | Out-Null
}

# Copy runtime files only when building from a separate source tree.
if ($srcDir -ne (Resolve-Path $distDir).Path) {
    # Clean previously copied runtime folders/files to avoid stale files in the package.
    foreach ($folder in $runtimeDirs) {
        $target = Join-Path $distDir $folder
        if (Test-Path $target) {
            Remove-Item -Path $target -Recurse -Force
        }
    }

    foreach ($file in $runtimeFiles) {
        $target = Join-Path $distDir $file
        if (Test-Path $target) {
            Remove-Item -Path $target -Force
        }
    }

    foreach ($folder in $runtimeDirs) {
        $source = Join-Path $srcDir $folder
        if (Test-Path $source) {
            Copy-Item -Path $source -Destination $distDir -Recurse -Force
        }
    }

    foreach ($file in $runtimeFiles) {
        $source = Join-Path $srcDir $file
        if (Test-Path $source) {
            Copy-Item -Path $source -Destination $distDir -Force
        }
    }
}

$out = Join-Path $distDir 'ADUserCreator.exe'
$icon = Join-Path $projectRoot 'assets\icon.ico'
$useIcon = Test-Path $icon

if ($useIcon) {
    ps2exe $srcMain $out -noConsole -x64 -requireAdmin -icon $icon -title 'AD User Creator' -verbose
} else {
    ps2exe $srcMain $out -noConsole -x64 -requireAdmin -title 'AD User Creator' -verbose
}

Write-Host "DONE: $out"
Write-Host "Runtime files copied to: $distDir"
