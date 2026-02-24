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
$releaseDir = Join-Path $projectRoot 'release'
$installerScript = Join-Path $PSScriptRoot 'ADUserCreator.iss'

if ($useIcon) {
    ps2exe $srcMain $out -noConsole -x64 -requireAdmin -icon $icon -title 'AD User Creator' -verbose
} else {
    ps2exe $srcMain $out -noConsole -x64 -requireAdmin -title 'AD User Creator' -verbose
}

if (-not (Test-Path $installerScript)) {
    throw "Не знайдено Inno Setup script: $installerScript"
}

if (-not (Test-Path $releaseDir)) {
    New-Item -ItemType Directory -Path $releaseDir | Out-Null
}

$requiredForInstaller = @(
    (Join-Path $distDir 'ADUserCreator.exe'),
    (Join-Path $distDir 'main.ps1'),
    (Join-Path $distDir 'ad'),
    (Join-Path $distDir 'common'),
    (Join-Path $distDir 'excel'),
    (Join-Path $distDir 'ui')
)

foreach ($path in $requiredForInstaller) {
    if (-not (Test-Path $path)) {
        throw "Для інсталятора відсутній файл/папка: $path"
    }
}

$isccCandidates = @()
$isccCommand = Get-Command ISCC.exe -ErrorAction SilentlyContinue
if ($isccCommand) {
    $isccCandidates += $isccCommand.Source
}

foreach ($candidate in @(
    (Join-Path ${env:ProgramFiles(x86)} 'Inno Setup 6\ISCC.exe'),
    (Join-Path $env:ProgramFiles 'Inno Setup 6\ISCC.exe')
)) {
    if ($candidate -and (Test-Path $candidate)) {
        $isccCandidates += $candidate
    }
}

$isccExe = $isccCandidates | Select-Object -Unique | Select-Object -First 1
if (-not $isccExe) {
    throw "Не знайдено ISCC.exe (Inno Setup 6). Встановіть Inno Setup і повторіть збірку."
}

$appVersion = Get-Date -Format 'yyyy.MM.dd.HHmm'
$isccArgs = @(
    "/DSourceDist=`"$distDir`"",
    "/DOutputDir=`"$releaseDir`"",
    "/DAppVersion=`"$appVersion`"",
    "`"$installerScript`""
)

& $isccExe @isccArgs
if ($LASTEXITCODE -ne 0) {
    throw "ISCC.exe завершився з кодом $LASTEXITCODE"
}

Write-Host "DONE: $out"
Write-Host "Runtime files copied to: $distDir"
Write-Host "Installer created in: $releaseDir"
