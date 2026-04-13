function ConvertTo-HtmlEncoded {
    param([string]$Value)
    return [System.Net.WebUtility]::HtmlEncode((Normalize-Text $Value))
}

function Get-PdfRendererPath {
    $candidates = @(
        'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
        'C:\Program Files\Microsoft\Edge\Application\msedge.exe',
        'C:\Program Files\Google\Chrome\Application\chrome.exe'
    )
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) { return $candidate }
    }
    throw 'Microsoft Edge or Google Chrome was not found for PDF generation.'
}

function Convert-HtmlFileToPdf {
    param(
        [Parameter(Mandatory)][string]$HtmlPath,
        [Parameter(Mandatory)][string]$PdfPath
    )

    $browserPath = Get-PdfRendererPath
    $htmlUri = [System.Uri]::new($HtmlPath).AbsoluteUri
    $pdfDir = Split-Path -Parent $PdfPath
    if (-not (Test-Path -LiteralPath $pdfDir)) {
        New-Item -ItemType Directory -Path $pdfDir -Force | Out-Null
    }

    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('adusercreator-pdf-' + [guid]::NewGuid().ToString('N'))
    $userDataDir = Join-Path $tempRoot 'profile'
    $stdoutPath = Join-Path $tempRoot 'stdout.log'
    $stderrPath = Join-Path $tempRoot 'stderr.log'
    New-Item -ItemType Directory -Path $userDataDir -Force | Out-Null

    $headlessModes = @('new', 'old')
    $lastError = $null

    try {
        foreach ($mode in $headlessModes) {
            if (Test-Path -LiteralPath $PdfPath) {
                Remove-Item -LiteralPath $PdfPath -Force -ErrorAction SilentlyContinue
            }
            if (Test-Path -LiteralPath $stdoutPath) {
                Remove-Item -LiteralPath $stdoutPath -Force -ErrorAction SilentlyContinue
            }
            if (Test-Path -LiteralPath $stderrPath) {
                Remove-Item -LiteralPath $stderrPath -Force -ErrorAction SilentlyContinue
            }

            $arguments = @(
                "--headless=$mode",
                '--disable-gpu',
                '--no-first-run',
                '--no-default-browser-check',
                '--allow-file-access-from-files',
                "--user-data-dir=$userDataDir",
                "--print-to-pdf=$PdfPath",
                $htmlUri
            )

            $process = Start-Process -FilePath $browserPath -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
            if ($process.ExitCode -eq 0) {
                for ($attempt = 0; $attempt -lt 10; $attempt++) {
                    if ((Test-Path -LiteralPath $PdfPath) -and ((Get-Item -LiteralPath $PdfPath).Length -gt 0)) {
                        return
                    }
                    Start-Sleep -Milliseconds 300
                }
            }

            $stderrText = if (Test-Path -LiteralPath $stderrPath) { (Get-Content -LiteralPath $stderrPath -Raw -ErrorAction SilentlyContinue).Trim() } else { '' }
            $stdoutText = if (Test-Path -LiteralPath $stdoutPath) { (Get-Content -LiteralPath $stdoutPath -Raw -ErrorAction SilentlyContinue).Trim() } else { '' }
            $details = @()
            if ($stderrText) { $details += "stderr: $stderrText" }
            if ($stdoutText) { $details += "stdout: $stdoutText" }
            $detailSuffix = if ($details.Count -gt 0) { ' ' + ($details -join ' | ') } else { '' }
            $lastError = "PDF generation failed in headless mode '$mode'. Exit code: $($process.ExitCode).$detailSuffix"
        }
    } finally {
        if (Test-Path -LiteralPath $tempRoot) {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    if ($lastError) {
        throw $lastError
    }
    throw 'PDF file was not created by the browser.'
}

function New-PasswordLogHtml {
    param(
        [Parameter(Mandatory)][object[]]$Rows,
        [Parameter(Mandatory)][string]$DomainSuffix,
        [Parameter(Mandatory)][string]$OU,
        [Parameter(Mandatory)][datetime]$GeneratedAt
    )

    $tickets = foreach ($row in $Rows) {
        $fullName = ConvertTo-HtmlEncoded $row.fullName
        $unit = ConvertTo-HtmlEncoded $row.unit
        $login = ConvertTo-HtmlEncoded $row.login
        $password = ConvertTo-HtmlEncoded $row.password
        @"
<section class="ticket">
  <div class="ticket-card">
    <div class="person-name">$fullName</div>
    <div class="person-unit">$unit</div>
    <div class="credential-row">&#1051;&#1086;&#1075;&#1110;&#1085;: $login</div>
    <div class="credential-row">&#1055;&#1072;&#1088;&#1086;&#1083;&#1100;: $password</div>
  </div>
</section>
"@
    }

    $headerDomain = ConvertTo-HtmlEncoded $DomainSuffix
    $headerOu = ConvertTo-HtmlEncoded $OU
    $headerDate = ConvertTo-HtmlEncoded ($GeneratedAt.ToString('yyyy-MM-dd HH:mm:ss'))

    return @"
<!doctype html>
<html lang="uk">
<head>
  <meta charset="utf-8">
  <title>Password credentials</title>
  <style>
    @page { size: A4 landscape; margin: 0; }
    * { box-sizing: border-box; }
    html, body { margin: 0; padding: 0; font-family: "Segoe UI", Arial, sans-serif; background: #ffffff; color: #000000; }
    body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    .ticket {
      width: 100vw;
      height: 100vh;
      min-height: 210mm;
      display: flex;
      align-items: center;
      justify-content: center;
      page-break-after: always;
    }
    .ticket:last-child { page-break-after: auto; }
    .ticket-card {
      width: 78%;
      max-width: 980px;
      text-align: center;
      padding: 24px 28px;
    }
    .person-name {
      font-size: 22px;
      line-height: 1.3;
      font-weight: 700;
      margin-bottom: 4px;
    }
    .person-unit {
      font-size: 19px;
      line-height: 1.35;
      margin-bottom: 2px;
    }
    .credential-row {
      font-size: 19px;
      line-height: 1.4;
    }
    .meta {
      position: fixed;
      top: 10mm;
      left: 12mm;
      font-size: 10px;
      color: #666;
    }
  </style>
</head>
<body>
  <div class="meta">Domain: $headerDomain | OU: $headerOu | Generated: $headerDate</div>
  $($tickets -join "`n")
</body>
</html>
"@
}

function Save-PasswordLogPdf {
    param(
        [Parameter(Mandatory)][object[]]$CreatedRows,
        [Parameter(Mandatory)][string]$DomainSuffix,
        [Parameter(Mandatory)][string]$OU,
        [string]$OutputDirectory
    )

    $targetRoot = if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
        Ensure-PasswordLogsRoot
        $script:PasswordLogsRoot
    } else {
        $resolvedOutput = $OutputDirectory.Trim()
        if (-not (Test-Path -LiteralPath $resolvedOutput)) {
            New-Item -ItemType Directory -Path $resolvedOutput -Force | Out-Null
        }
        $resolvedOutput
    }

    $rowsWithPasswords = @($CreatedRows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.password) })
    if (-not $rowsWithPasswords.Count) { return $null }

    $stamp = Get-Date
    $id = $stamp.ToString('yyyyMMdd-HHmmss')
    $safeDomain = (($DomainSuffix -replace '[^a-zA-Z0-9._-]', '_').Trim('_'))
    if ([string]::IsNullOrWhiteSpace($safeDomain)) { $safeDomain = 'domain' }
    $fileName = "passwords_credentials_${safeDomain}_$id.pdf"
    $pdfPath = Join-Path $targetRoot $fileName
    $metaPath = [System.IO.Path]::ChangeExtension($pdfPath, '.json')

    $htmlPath = [System.IO.Path]::ChangeExtension($pdfPath, '.html')
    $htmlContent = New-PasswordLogHtml -Rows $rowsWithPasswords -DomainSuffix $DomainSuffix -OU $OU -GeneratedAt $stamp
    Set-Content -LiteralPath $htmlPath -Value $htmlContent -Encoding UTF8
    try {
        Convert-HtmlFileToPdf -HtmlPath $htmlPath -PdfPath $pdfPath
    } finally {
        if (Test-Path -LiteralPath $htmlPath) { Remove-Item -LiteralPath $htmlPath -Force -ErrorAction SilentlyContinue }
    }
    $pdfInfo = Get-Item -LiteralPath $pdfPath

    $meta = [pscustomobject]@{
        id = $id
        name = $fileName
        path = $pdfPath
        date = $stamp.ToString('yyyy-MM-dd HH:mm')
        createdAt = $stamp.ToString('o')
        users = $rowsWithPasswords.Count
        size = '{0:N0} KB' -f ([Math]::Max(1, [Math]::Round($pdfInfo.Length / 1KB, 0)))
        domain = $DomainSuffix
        ou = $OU
    }
    ($meta | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $metaPath -Encoding UTF8
    return $meta
}

function Get-PasswordLogItems {
    Ensure-PasswordLogsRoot
    $items = @(Get-ChildItem -LiteralPath $script:PasswordLogsRoot -Filter *.json -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        ForEach-Object {
            try {
                $meta = Get-Content -LiteralPath $_.FullName -Raw | ConvertFrom-Json
                if ($meta.name -and (Test-Path -LiteralPath $meta.path)) {
                    [pscustomobject]@{
                        id = [string]$meta.id
                        name = [string]$meta.name
                        date = [string]$meta.date
                        users = [int]$meta.users
                        size = [string]$meta.size
                        domain = [string]$meta.domain
                        ou = [string]$meta.ou
                    }
                }
            } catch {}
        })
    return $items
}

function Get-PasswordLogMetaById {
    param([Parameter(Mandatory)][string]$Id)
    Ensure-PasswordLogsRoot
    $metaFile = Get-ChildItem -LiteralPath $script:PasswordLogsRoot -Filter *.json -File -ErrorAction SilentlyContinue |
        Where-Object {
            try {
                $meta = Get-Content -LiteralPath $_.FullName -Raw | ConvertFrom-Json
                [string]$meta.id -eq $Id
            } catch {
                $false
            }
        } |
        Select-Object -First 1
    if (-not $metaFile) { throw 'PDF log was not found.' }
    return (Get-Content -LiteralPath $metaFile.FullName -Raw | ConvertFrom-Json)
}
