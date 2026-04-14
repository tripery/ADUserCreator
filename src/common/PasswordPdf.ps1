function Get-PdfRendererPath {
    $candidates = @(
        'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
        'C:\Program Files\Microsoft\Edge\Application\msedge.exe',
        'C:\Program Files\Google\Chrome\Application\chrome.exe'
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    throw 'Microsoft Edge or Google Chrome was not found for PDF generation.'
}

function Get-DefaultPasswordPdfDirectory {
    $downloads = Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads'
    if (-not (Test-Path -LiteralPath $downloads)) {
        New-Item -ItemType Directory -Path $downloads -Force | Out-Null
    }
    return $downloads
}

function ConvertTo-HtmlSafeText {
    param([string]$Value)
    return [System.Net.WebUtility]::HtmlEncode(([string]$Value).Trim())
}

function New-PasswordPdfHtml {
    param(
        [Parameter(Mandatory = $true)][object[]]$Rows,
        [Parameter(Mandatory = $true)][string]$DomainSuffix,
        [Parameter(Mandatory = $true)][string]$OU,
        [Parameter(Mandatory = $true)][datetime]$GeneratedAt
    )

    $tickets = foreach ($row in $Rows) {
        $displayName = ConvertTo-HtmlSafeText $row.DisplayName
        $login = ConvertTo-HtmlSafeText $row.SamAccountName
        $mail = ConvertTo-HtmlSafeText $row.Mail
        $password = ConvertTo-HtmlSafeText $row.Password
@"
<section class="ticket">
  <div class="card">
    <div class="name">$displayName</div>
    <div class="row">Login: $login</div>
    <div class="row">Mail: $mail</div>
    <div class="row password">Password: $password</div>
  </div>
</section>
"@
    }

    $domain = ConvertTo-HtmlSafeText $DomainSuffix
    $ouSafe = ConvertTo-HtmlSafeText $OU
    $generated = ConvertTo-HtmlSafeText ($GeneratedAt.ToString('yyyy-MM-dd HH:mm:ss'))

@"
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Password credentials</title>
  <style>
    @page { size: A4 landscape; margin: 0; }
    * { box-sizing: border-box; }
    html, body { margin: 0; padding: 0; font-family: "Segoe UI", Arial, sans-serif; background: #fff; color: #111; }
    body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    .meta { position: fixed; top: 10mm; left: 12mm; font-size: 10px; color: #666; }
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
    .card {
      width: 72%;
      max-width: 920px;
      padding: 28px 32px;
      text-align: center;
    }
    .name {
      font-size: 24px;
      font-weight: 700;
      margin-bottom: 14px;
    }
    .row {
      font-size: 19px;
      line-height: 1.45;
    }
    .password {
      margin-top: 4px;
      font-weight: 700;
    }
  </style>
</head>
<body>
  <div class="meta">Domain: $domain | OU: $ouSafe | Generated: $generated</div>
  $($tickets -join "`n")
</body>
</html>
"@
}

function Convert-HtmlFileToPdf {
    param(
        [Parameter(Mandatory = $true)][string]$HtmlPath,
        [Parameter(Mandatory = $true)][string]$PdfPath
    )

    $browserPath = Get-PdfRendererPath
    $htmlUri = [System.Uri]::new($HtmlPath).AbsoluteUri
    $pdfDir = Split-Path -Parent $PdfPath
    if (-not (Test-Path -LiteralPath $pdfDir)) {
        New-Item -ItemType Directory -Path $pdfDir -Force | Out-Null
    }

    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('adusercreator-pdf-' + [guid]::NewGuid().ToString('N'))
    $userDataDir = Join-Path $tempRoot 'profile'
    New-Item -ItemType Directory -Path $userDataDir -Force | Out-Null

    try {
        $arguments = @(
            '--headless=new',
            '--disable-gpu',
            '--no-first-run',
            '--no-default-browser-check',
            '--allow-file-access-from-files',
            "--user-data-dir=$userDataDir",
            "--print-to-pdf=$PdfPath",
            $htmlUri
        )

        $process = Start-Process -FilePath $browserPath -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden
        if ($process.ExitCode -ne 0) {
            throw "PDF generation failed. Browser exit code: $($process.ExitCode)"
        }
    }
    finally {
        if (Test-Path -LiteralPath $tempRoot) {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    if (-not (Test-Path -LiteralPath $PdfPath)) {
        throw 'PDF file was not created.'
    }
}

function Save-PasswordCredentialsPdfToFolder {
    param(
        [Parameter(Mandatory = $true)][object[]]$Rows,
        [Parameter(Mandatory = $true)][string]$DomainSuffix,
        [Parameter(Mandatory = $true)][string]$OU,
        [Parameter(Mandatory = $true)][string]$OutputDirectory
    )

    $rowsWithPasswords = @($Rows | Where-Object { $_.Status -eq 'OK' -and -not [string]::IsNullOrWhiteSpace($_.Password) })
    if (-not $rowsWithPasswords.Count) {
        return $null
    }

    $targetDirectory = $OutputDirectory.Trim()
    if ([string]::IsNullOrWhiteSpace($targetDirectory)) {
        $targetDirectory = Get-DefaultPasswordPdfDirectory
    }
    if (-not (Test-Path -LiteralPath $targetDirectory)) {
        New-Item -ItemType Directory -Path $targetDirectory -Force | Out-Null
    }

    $safeDomain = (($DomainSuffix -replace '[^a-zA-Z0-9._-]', '_').Trim('_'))
    if ([string]::IsNullOrWhiteSpace($safeDomain)) { $safeDomain = 'domain' }

    $stamp = Get-Date
    $pdfPath = Join-Path $targetDirectory ("passwords_credentials_{0}_{1}.pdf" -f $safeDomain, $stamp.ToString('yyyyMMdd-HHmmss'))
    $htmlPath = Join-Path ([System.IO.Path]::GetTempPath()) ("adusercreator-passwords-{0}.html" -f ([guid]::NewGuid().ToString('N')))

    $html = New-PasswordPdfHtml -Rows $rowsWithPasswords -DomainSuffix $DomainSuffix -OU $OU -GeneratedAt $stamp
    Set-Content -LiteralPath $htmlPath -Value $html -Encoding UTF8

    try {
        Convert-HtmlFileToPdf -HtmlPath $htmlPath -PdfPath $pdfPath
    }
    finally {
        if (Test-Path -LiteralPath $htmlPath) {
            Remove-Item -LiteralPath $htmlPath -Force -ErrorAction SilentlyContinue
        }
    }

    return $pdfPath
}
