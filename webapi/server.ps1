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
$script:PasswordLogsRoot = Join-Path $script:ProjectRoot 'data\password-logs'

function Ensure-PasswordLogsRoot {
    if (-not (Test-Path -LiteralPath $script:PasswordLogsRoot)) {
        New-Item -ItemType Directory -Path $script:PasswordLogsRoot -Force | Out-Null
    }
}

function Initialize-AppDependencies {
    Import-Module ActiveDirectory -ErrorAction Stop
}

. (Join-Path $script:ProjectRoot 'src\core\ad\Transliteration.ps1')
. (Join-Path $script:ProjectRoot 'src\core\ad\Naming.ps1')
. (Join-Path $script:ProjectRoot 'src\core\ad\UserPreview.ps1')
. (Join-Path $script:ProjectRoot 'src\core\ad\UserCreate.ps1')
. (Join-Path $script:ProjectRoot 'src\core\common\Password.ps1')
. (Join-Path $script:ProjectRoot 'src\web\Http.ps1')
. (Join-Path $script:ProjectRoot 'src\logs\PasswordLogs.ps1')

function Normalize-Text {
    param([string]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Trim()
}

function Get-AdOuOptions {
    $domain = Get-ADDomain
    $domainDN = $domain.DistinguishedName
    $allObjects = Get-ADObject -LDAPFilter '(|(objectClass=organizationalUnit)(objectClass=container))' -SearchBase $domainDN -ErrorAction Stop | Sort-Object DistinguishedName
    [pscustomobject]@{
        domainDnsRoot = $domain.DNSRoot
        domainDN      = $domainDN
        items         = @(
            $allObjects | ForEach-Object {
                [pscustomobject]@{
                    name              = $_.Name
                    distinguishedName = $_.DistinguishedName
                }
            }
        )
    }
}

function Get-AdGroupOptions {
    @(
        Get-ADGroup -Filter { GroupCategory -eq 'Security' } -Properties Name, SamAccountName, GroupScope |
            Sort-Object Name |
            ForEach-Object {
                [pscustomobject]@{
                    name           = $_.Name
                    samAccountName = $_.SamAccountName
                    scope          = [string]$_.GroupScope
                }
            }
    )
}

function Handle-ApiRequest {
    param([Parameter(Mandatory)]$Context)

    $req = $Context.Request
    $method = $req.HttpMethod.ToUpperInvariant()
    $path = $req.Url.AbsolutePath.TrimEnd('/')
    if ([string]::IsNullOrWhiteSpace($path)) { $path = '/' }

    if ($method -eq 'OPTIONS') {
        Write-TextResponse -Context $Context -Text '' -StatusCode 204
        return
    }

    try {
        switch ("$method $path") {
            'GET /api/health' {
                Write-JsonResponse -Context $Context -Data ([pscustomobject]@{
                        ok         = $true
                        serverTime = (Get-Date).ToString('s')
                        machine    = $env:COMPUTERNAME
                        user       = $env:USERNAME
                    })
                return
            }
            'GET /api/ad/options' {
                $ouData = Get-AdOuOptions
                $groups = Get-AdGroupOptions
                Write-JsonResponse -Context $Context -Data ([pscustomobject]@{
                        ok       = $true
                        domain   = $ouData.domainDnsRoot
                        domainDN = $ouData.domainDN
                        ous      = $ouData.items
                        groups   = $groups
                    })
                return
            }
            'GET /api/password-logs' {
                Write-JsonResponse -Context $Context -Data ([pscustomobject]@{
                        ok    = $true
                        items = @(Get-PasswordLogItems)
                    })
                return
            }
            'GET /api/password-logs/file' {
                $id = Normalize-Text $req.QueryString['id']
                if ([string]::IsNullOrWhiteSpace($id)) { throw 'id є обов''язковим.' }

                $downloadRaw = Normalize-Text $req.QueryString['download']
                $download = $false
                if (-not [string]::IsNullOrWhiteSpace($downloadRaw)) {
                    $download = [System.Convert]::ToBoolean($downloadRaw)
                }

                $meta = Get-PasswordLogMetaById -Id $id
                if (-not (Test-Path -LiteralPath ([string]$meta.path))) {
                    throw 'PDF файл на диску не знайдено.'
                }

                $pdfBytes = [System.IO.File]::ReadAllBytes([string]$meta.path)
                Write-BinaryResponse -Context $Context -Bytes $pdfBytes -ContentType 'application/pdf' -FileName ([string]$meta.name) -Download:$download
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
                    try {
                        $preview.Add((New-PreviewUserRecord -UserItem $u -DomainSuffix $domainSuffix -OU $ou -CheckUniqueness))
                    } catch {
                        $errors.Add([pscustomobject]@{
                                fullName  = (Normalize-Text $u.fullName)
                                sourceRow = $u.sourceRow
                                error     = $_.Exception.Message
                            })
                    }
                }

                Write-JsonResponse -Context $Context -Data ([pscustomobject]@{
                        ok      = $true
                        preview = $preview.ToArray()
                        errors  = $errors.ToArray()
                    })
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
                            $results.Add([pscustomobject]@{
                                    fullName = $preview.fullName
                                    login    = $preview.login
                                    email    = $preview.email
                                    status   = 'dry-run'
                                })
                        } else {
                            $results.Add((Invoke-AdUserCreate -Preview $preview -OU $ou -GroupsToAdd $groupsToAdd -PasswordNeverExpires $passwordNeverExpires))
                        }
                    } catch {
                        $errors.Add([pscustomobject]@{
                                fullName  = (Normalize-Text $u.fullName)
                                sourceRow = $u.sourceRow
                                error     = $_.Exception.Message
                            })
                    }
                }

                $pdfLog = $null
                if (-not $dryRun -and $results.Count -gt 0) {
                    $pdfLog = Save-PasswordLogPdf -CreatedRows $results.ToArray() -DomainSuffix $domainSuffix -OU $ou
                }

                Write-JsonResponse -Context $Context -Data ([pscustomobject]@{
                        ok           = $true
                        created      = $results.ToArray()
                        errors       = $errors.ToArray()
                        pdfLog       = $pdfLog
                        passwordLogs = @(Get-PasswordLogItems)
                    })
                return
            }
            default {
                Write-JsonResponse -Context $Context -Data ([pscustomobject]@{
                        ok    = $false
                        error = 'Not found'
                    }) -StatusCode 404
                return
            }
        }
    } catch {
        Write-JsonResponse -Context $Context -Data ([pscustomobject]@{
                ok    = $false
                error = $_.Exception.Message
            }) -StatusCode 500
    }
}

Initialize-AppDependencies
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://+:$Port/")
try {
    $listener.Start()
} catch [System.Net.HttpListenerException] {
    $currentUser = "$env:USERDOMAIN\$env:USERNAME"
    $urlAclCmd = "netsh http add urlacl url=http://+:$Port/ user=`"$currentUser`""
    throw "Не вдалося запустити HttpListener на http://+:$Port/. Найчастіше причина: URL ACL не налаштований. Запустіть PowerShell від адміністратора і виконайте: $urlAclCmd"
}

Write-Host "ADUserCreator Web API started on http://localhost:$Port/api/health"
Write-Host "LAN access: http://10.21.2.105:$Port/api/health"
Write-Host 'Press Ctrl+C to stop'

try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        Handle-ApiRequest -Context $context
    }
} finally {
    if ($listener.IsListening) { $listener.Stop() }
    $listener.Close()
}
