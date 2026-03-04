param(
    [int]$Port = 8787,
    [string]$AllowOrigin = '*'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Resolve-ProjectRoot {
    if ($PSScriptRoot) { return (Split-Path $PSScriptRoot -Parent) }
    return (Get-Location).Path
}

$script:ProjectRoot = Resolve-ProjectRoot

function Initialize-AppDependencies {
    Import-Module ActiveDirectory -ErrorAction Stop
    . (Join-Path $script:ProjectRoot 'src\ad\Transliteration.ps1')
    . (Join-Path $script:ProjectRoot 'src\ad\Naming.ps1')
    . (Join-Path $script:ProjectRoot 'src\common\Password.ps1')
}

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

function Read-JsonBody {
    param([Parameter(Mandatory)]$Request)
    $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
    try { $body = $reader.ReadToEnd() } finally { $reader.Dispose() }
    if ([string]::IsNullOrWhiteSpace($body)) { return @{} }
    return ($body | ConvertFrom-Json -Depth 12)
}

function Normalize-Text {
    param([string]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Trim()
}

function Split-FullName {
    param([Parameter(Mandatory)][string]$FullName)
    $parts = (Normalize-Text $FullName) -split '\s+' | Where-Object { $_ }
    if ($parts.Count -lt 2) { throw "Неможливо розібрати ПІБ '$FullName'. Очікується щонайменше 'Прізвище Ім''я'." }
    [pscustomobject]@{
        Surname    = $parts[0]
        GivenName  = $parts[1]
        MiddleName = if ($parts.Count -ge 3) { ($parts[2..($parts.Count - 1)] -join ' ') } else { '' }
    }
}

function Build-BaseIdentifiers {
    param([Parameter(Mandatory)]$NameParts)
    $surnameLat = Convert-UA2Latin $NameParts.Surname
    $givenLat = Convert-UA2Latin $NameParts.GivenName
    $middleLat = if ($NameParts.MiddleName) { Convert-UA2Latin $NameParts.MiddleName } else { '' }
    if ([string]::IsNullOrWhiteSpace($surnameLat) -or [string]::IsNullOrWhiteSpace($givenLat)) { throw 'Не вдалося транслітерувати ПІБ у латиницю.' }
    $samBase = ('{0}.{1}' -f $givenLat.Substring(0,1), $surnameLat).ToLower()
    [pscustomobject]@{ SurnameLatin = $surnameLat; GivenLatin = $givenLat; MiddleLatin = $middleLat; SamBase = $samBase; MailLocalBase = $samBase }
}

function New-PreviewUserRecord {
    param([Parameter(Mandatory)]$UserItem,[Parameter(Mandatory)][string]$DomainSuffix,[string]$OU,[switch]$CheckUniqueness)
    $fullName = Normalize-Text $UserItem.fullName
    if ([string]::IsNullOrWhiteSpace($fullName)) { throw 'Порожнє поле fullName у записі користувача.' }
    $parts = Split-FullName -FullName $fullName
    $ids = Build-BaseIdentifiers -NameParts $parts
    $cn = $fullName; $sam = $ids.SamBase; $mailLocal = $ids.MailLocalBase
    if ($CheckUniqueness) {
        $sam = Get-UniqueSamAccountName -BaseSam $sam
        $mailLocal = Get-UniqueMailLocalPart -BaseLocal $mailLocal -DomainSuffix $DomainSuffix
        if (-not [string]::IsNullOrWhiteSpace($OU)) { $cn = Get-UniqueCN -BaseCN $cn -OUPath $OU }
    }
    [pscustomobject]@{
        fullName = $fullName
        surname = $parts.Surname
        givenName = $parts.GivenName
        middleName = $parts.MiddleName
        login = $sam
        email = "$mailLocal@$DomainSuffix"
        upn = "$sam@$DomainSuffix"
        cn = $cn
        unit = (Normalize-Text $UserItem.unit)
        sourceRow = $UserItem.sourceRow
    }
}

function Invoke-AdUserCreate {
    param([Parameter(Mandatory)]$Preview,[Parameter(Mandatory)][string]$OU,[Parameter(Mandatory)][string[]]$GroupsToAdd,[bool]$PasswordNeverExpires = $false)
    $password = Get-RandomPassword
    $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
    $newParams = @{
        Name = $Preview.cn
        DisplayName = $Preview.fullName
        GivenName = $Preview.givenName
        Surname = $Preview.surname
        SamAccountName = $Preview.login
        UserPrincipalName = $Preview.upn
        EmailAddress = $Preview.email
        Path = $OU
        Enabled = $true
        AccountPassword = $securePassword
        ChangePasswordAtLogon = $true
    }
    if ($Preview.middleName) { $newParams['OtherName'] = $Preview.middleName }
    if ($PasswordNeverExpires) { $newParams['PasswordNeverExpires'] = $true }
    New-ADUser @newParams
    foreach ($groupSam in $GroupsToAdd) { if ($groupSam) { Add-ADGroupMember -Identity $groupSam -Members $Preview.login -ErrorAction Stop } }
    [pscustomobject]@{ fullName = $Preview.fullName; login = $Preview.login; email = $Preview.email; password = $password; status = 'created' }
}

function Get-AdOuOptions {
    $domain = Get-ADDomain
    $domainDN = $domain.DistinguishedName
    $allObjects = Get-ADObject -LDAPFilter '(|(objectClass=organizationalUnit)(objectClass=container))' -SearchBase $domainDN -ErrorAction Stop | Sort-Object DistinguishedName
    [pscustomobject]@{
        domainDnsRoot = $domain.DNSRoot
        domainDN = $domainDN
        items = @($allObjects | ForEach-Object { [pscustomobject]@{ name = $_.Name; distinguishedName = $_.DistinguishedName } })
    }
}

function Get-AdGroupOptions {
    @(Get-ADGroup -Filter { GroupCategory -eq 'Security' } -Properties Name,SamAccountName,GroupScope |
      Sort-Object Name |
      ForEach-Object { [pscustomobject]@{ name = $_.Name; samAccountName = $_.SamAccountName; scope = [string]$_.GroupScope } })
}

function Handle-ApiRequest {
    param([Parameter(Mandatory)]$Context)
    $req = $Context.Request
    $method = $req.HttpMethod.ToUpperInvariant()
    $path = $req.Url.AbsolutePath.TrimEnd('/')
    if ([string]::IsNullOrWhiteSpace($path)) { $path = '/' }

    if ($method -eq 'OPTIONS') { Write-TextResponse -Context $Context -Text '' -StatusCode 204; return }

    try {
        switch ("$method $path") {
            'GET /api/health' {
                Write-JsonResponse -Context $Context -Data ([pscustomobject]@{ ok = $true; serverTime = (Get-Date).ToString('s'); machine = $env:COMPUTERNAME; user = $env:USERNAME })
                return
            }
            'GET /api/ad/options' {
                $ouData = Get-AdOuOptions
                $groups = Get-AdGroupOptions
                Write-JsonResponse -Context $Context -Data ([pscustomobject]@{ ok = $true; domain = $ouData.domainDnsRoot; domainDN = $ouData.domainDN; ous = $ouData.items; groups = $groups })
                return
            }
            'POST /api/users/preview' {
                $body = Read-JsonBody -Request $req
                $domainSuffix = Normalize-Text $body.domainSuffix
                if ([string]::IsNullOrWhiteSpace($domainSuffix)) { throw 'domainSuffix є обов''язковим.' }
                $users = @($body.users)
                $ou = Normalize-Text $body.ou
                $preview = New-Object System.Collections.Generic.List[object]
                $errors = New-Object System.Collections.Generic.List[object]
                foreach ($u in $users) {
                    try { $preview.Add((New-PreviewUserRecord -UserItem $u -DomainSuffix $domainSuffix -OU $ou -CheckUniqueness)) }
                    catch { $errors.Add([pscustomobject]@{ fullName = (Normalize-Text $u.fullName); sourceRow = $u.sourceRow; error = $_.Exception.Message }) }
                }
                Write-JsonResponse -Context $Context -Data ([pscustomobject]@{ ok = $true; preview = @($preview); errors = @($errors) })
                return
            }
            'POST /api/users/create' {
                $body = Read-JsonBody -Request $req
                $domainSuffix = Normalize-Text $body.domainSuffix
                $ou = Normalize-Text $body.ou
                if ([string]::IsNullOrWhiteSpace($domainSuffix)) { throw 'domainSuffix є обов''язковим.' }
                if ([string]::IsNullOrWhiteSpace($ou)) { throw 'ou є обов''язковим.' }
                $groupsToAdd = @($body.groupsToAdd | ForEach-Object { Normalize-Text $_ } | Where-Object { $_ })
                $passwordNeverExpires = [bool]$body.passwordNeverExpires
                $dryRun = [bool]$body.dryRun
                $users = @($body.users)
                $results = New-Object System.Collections.Generic.List[object]
                $errors = New-Object System.Collections.Generic.List[object]
                foreach ($u in $users) {
                    try {
                        $preview = New-PreviewUserRecord -UserItem $u -DomainSuffix $domainSuffix -OU $ou -CheckUniqueness
                        if ($dryRun) {
                            $results.Add([pscustomobject]@{ fullName = $preview.fullName; login = $preview.login; email = $preview.email; status = 'dry-run' })
                        } else {
                            $results.Add((Invoke-AdUserCreate -Preview $preview -OU $ou -GroupsToAdd $groupsToAdd -PasswordNeverExpires $passwordNeverExpires))
                        }
                    } catch {
                        $errors.Add([pscustomobject]@{ fullName = (Normalize-Text $u.fullName); sourceRow = $u.sourceRow; error = $_.Exception.Message })
                    }
                }
                Write-JsonResponse -Context $Context -Data ([pscustomobject]@{ ok = $true; created = @($results); errors = @($errors) })
                return
            }
            default {
                Write-JsonResponse -Context $Context -Data ([pscustomobject]@{ ok = $false; error = 'Not found' }) -StatusCode 404
                return
            }
        }
    } catch {
        Write-JsonResponse -Context $Context -Data ([pscustomobject]@{ ok = $false; error = $_.Exception.Message }) -StatusCode 500
    }
}

Initialize-AppDependencies
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$Port/")
$listener.Prefixes.Add("http://127.0.0.1:$Port/")
$listener.Start()
Write-Host "ADUserCreator Web API started on http://localhost:$Port/api/health"
Write-Host 'Press Ctrl+C to stop'
Write-Host 'NOTE: src/ad/UserProvision.ps1 is currently broken in repo; backend uses inline create logic.'
try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        Handle-ApiRequest -Context $context
    }
} finally {
    if ($listener.IsListening) { $listener.Stop() }
    $listener.Close()
}
