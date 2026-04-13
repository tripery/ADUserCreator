function ConvertTo-HtmlEncoded {
    param([string]$Value)
    return [System.Net.WebUtility]::HtmlEncode(([string]$Value).Trim())
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
        [Parameter(Mandatory = $true)][string]$HtmlPath,
        [Parameter(Mandatory = $true)][string]$PdfPath
    )

    $browserPath = Get-PdfRendererPath
    $htmlUri = [System.Uri]::new($HtmlPath).AbsoluteUri
    $arguments = @(
        '--headless=new',
        '--disable-gpu',
        '--no-first-run',
        '--no-default-browser-check',
        "--print-to-pdf=$PdfPath",
        $htmlUri
    )

    $process = Start-Process -FilePath $browserPath -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden
    if ($process.ExitCode -ne 0) {
        throw "Failed to generate PDF. Browser exit code: $($process.ExitCode)"
    }

    if (-not (Test-Path -LiteralPath $PdfPath)) {
        throw 'PDF file was not created by the browser.'
    }
}

function New-PasswordLogHtml {
    param(
        [Parameter(Mandatory = $true)][object[]]$Rows,
        [Parameter(Mandatory = $true)][string]$DomainSuffix,
        [Parameter(Mandatory = $true)][string]$OU,
        [Parameter(Mandatory = $true)][datetime]$GeneratedAt
    )

    $tickets = foreach ($row in $Rows) {
        $fullName = ConvertTo-HtmlEncoded $row.DisplayName
        $unit = ConvertTo-HtmlEncoded $row.Department
        $login = ConvertTo-HtmlEncoded $row.SamAccountName
        $password = ConvertTo-HtmlEncoded $row.Password
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

@"
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

function Save-PasswordCredentialsPdf {
    param(
        [Parameter(Mandatory = $true)][object[]]$Rows,
        [Parameter(Mandatory = $true)][string]$DomainSuffix,
        [Parameter(Mandatory = $true)][string]$OU,
        [string]$PdfPath
    )

    $rowsWithPasswords = @($Rows | Where-Object {
        $_.Status -eq 'OK' -and -not [string]::IsNullOrWhiteSpace($_.Password)
    })
    if (-not $rowsWithPasswords.Count) { return $null }

    $stamp = Get-Date
    if ([string]::IsNullOrWhiteSpace($PdfPath)) {
        $safeDomain = (($DomainSuffix -replace '[^a-zA-Z0-9._-]', '_').Trim('_'))
        if ([string]::IsNullOrWhiteSpace($safeDomain)) { $safeDomain = 'domain' }
        $downloads = [Environment]::GetFolderPath('UserProfile')
        $downloads = Join-Path $downloads 'Downloads'
        $PdfPath = Join-Path $downloads ("passwords_credentials_{0}_{1}.pdf" -f $safeDomain, $stamp.ToString('yyyyMMdd-HHmmss'))
    }

    $htmlPath = Join-Path ([System.IO.Path]::GetTempPath()) ("adusercreator-passwords-{0}.html" -f ([guid]::NewGuid().ToString('N')))
    $htmlContent = New-PasswordLogHtml -Rows $rowsWithPasswords -DomainSuffix $DomainSuffix -OU $OU -GeneratedAt $stamp
    Set-Content -LiteralPath $htmlPath -Value $htmlContent -Encoding UTF8

    try {
        Convert-HtmlFileToPdf -HtmlPath $htmlPath -PdfPath $PdfPath
    }
    finally {
        if (Test-Path -LiteralPath $htmlPath) {
            Remove-Item -LiteralPath $htmlPath -Force -ErrorAction SilentlyContinue
        }
    }

    return $PdfPath
}

function Save-PasswordCredentialsPdfInteractive {
    param(
        [Parameter(Mandatory = $true)][object[]]$Rows,
        [Parameter(Mandatory = $true)][string]$DomainSuffix,
        [Parameter(Mandatory = $true)][string]$OU
    )

    $rowsWithPasswords = @($Rows | Where-Object {
        $_.Status -eq 'OK' -and -not [string]::IsNullOrWhiteSpace($_.Password)
    })
    if (-not $rowsWithPasswords.Count) { return $null }

    $safeDomain = (($DomainSuffix -replace '[^a-zA-Z0-9._-]', '_').Trim('_'))
    if ([string]::IsNullOrWhiteSpace($safeDomain)) { $safeDomain = 'domain' }

    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = 'PDF files (*.pdf)|*.pdf'
    $dlg.FileName = "passwords_credentials_{0}_{1}.pdf" -f $safeDomain, (Get-Date -Format 'yyyyMMdd-HHmmss')
    $dlg.InitialDirectory = Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads'

    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return $null
    }

    return (Save-PasswordCredentialsPdf -Rows $rowsWithPasswords -DomainSuffix $DomainSuffix -OU $OU -PdfPath $dlg.FileName)
}

function Get-DefaultPasswordPdfPath {
    param([string]$DomainSuffix)

    $safeDomain = (($DomainSuffix -replace '[^a-zA-Z0-9._-]', '_').Trim('_'))
    if ([string]::IsNullOrWhiteSpace($safeDomain)) { $safeDomain = 'domain' }

    $downloads = [Environment]::GetFolderPath('UserProfile')
    $downloads = Join-Path $downloads 'Downloads'
    return (Join-Path $downloads ("passwords_credentials_{0}_{1}.pdf" -f $safeDomain, (Get-Date -Format 'yyyyMMdd-HHmmss')))
}
