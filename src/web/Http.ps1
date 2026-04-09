function Add-CorsHeaders {
    param([Parameter(Mandatory)]$Response)
    $Response.Headers['Access-Control-Allow-Origin'] = $AllowOrigin
    $Response.Headers['Access-Control-Allow-Methods'] = 'GET,POST,OPTIONS'
    $Response.Headers['Access-Control-Allow-Headers'] = 'Content-Type'
}

function Write-JsonResponse {
    param([Parameter(Mandatory)]$Context,[Parameter(Mandatory)]$Data,[int]$StatusCode = 200)
    $json = $Data | ConvertTo-Json -Depth 12
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $res = $Context.Response
    $res.StatusCode = $StatusCode
    $res.ContentType = 'application/json; charset=utf-8'
    Add-CorsHeaders -Response $res
    $res.ContentEncoding = [System.Text.Encoding]::UTF8
    $res.OutputStream.Write($bytes, 0, $bytes.Length)
    $res.OutputStream.Close()
}

function Write-TextResponse {
    param([Parameter(Mandatory)]$Context,[Parameter(Mandatory)][string]$Text,[int]$StatusCode = 200,[string]$ContentType = 'text/plain; charset=utf-8')
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
    $res = $Context.Response
    $res.StatusCode = $StatusCode
    $res.ContentType = $ContentType
    Add-CorsHeaders -Response $res
    $res.ContentEncoding = [System.Text.Encoding]::UTF8
    $res.OutputStream.Write($bytes, 0, $bytes.Length)
    $res.OutputStream.Close()
}

function Write-BinaryResponse {
    param(
        [Parameter(Mandatory)]$Context,
        [Parameter(Mandatory)][byte[]]$Bytes,
        [int]$StatusCode = 200,
        [string]$ContentType = 'application/octet-stream',
        [string]$FileName,
        [switch]$Download
    )
    $res = $Context.Response
    $res.StatusCode = $StatusCode
    $res.ContentType = $ContentType
    $res.ContentLength64 = $Bytes.Length
    if ($FileName) {
        $disposition = if ($Download) { 'attachment' } else { 'inline' }
        $res.Headers['Content-Disposition'] = "$disposition; filename=""$FileName"""
    }
    Add-CorsHeaders -Response $res
    $res.OutputStream.Write($Bytes, 0, $Bytes.Length)
    $res.OutputStream.Close()
}

function Read-JsonBody {
    param([Parameter(Mandatory)]$Request)
    # Browser JSON payload is UTF-8; HttpListener ContentEncoding may be incorrect when charset is omitted.
    $reader = New-Object System.IO.StreamReader($Request.InputStream, [System.Text.Encoding]::UTF8)
    try { $body = $reader.ReadToEnd() } finally { $reader.Dispose() }
    if ([string]::IsNullOrWhiteSpace($body)) { return @{} }
    return ($body | ConvertFrom-Json)
}
