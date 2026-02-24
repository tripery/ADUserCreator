function Get-UniqueCN {
    param(
        [Parameter(Mandatory=$true)][string]$BaseCN,
        [Parameter(Mandatory=$true)][string]$OUPath
    )

    $cn = $BaseCN
    $i = 0
    do {
        $filter = "(&(cn=$cn)(objectClass=user))"
        $existing = Get-ADObject -LDAPFilter $filter -SearchBase $OUPath -ErrorAction SilentlyContinue
        if ($existing) { $i++; $cn = "$BaseCN ($i)" }
    } while ($existing)

    return $cn
}

function Get-UniqueSamAccountName {
    param([Parameter(Mandatory=$true)][string]$BaseSam)

    $base = $BaseSam.ToLower()
    if ($base.Length -gt 20) { $base = $base.Substring(0,20) }

    $sam = $base
    $i = 1
    while (Get-ADUser -Filter "SamAccountName -eq '$sam'" -ErrorAction SilentlyContinue) {
        $i++
        $suffix = $i.ToString()
        $maxLen = 20 - $suffix.Length
        $sam = ($base.Substring(0, [Math]::Min($base.Length, $maxLen))) + $suffix
    }
    return $sam
}

function Get-UniqueMailLocalPart {
    param(
        [Parameter(Mandatory=$true)][string]$BaseLocal,
        [Parameter(Mandatory=$true)][string]$DomainSuffix
    )

    $searchBase = (Get-ADDomain).DistinguishedName
    $base = $BaseLocal.ToLower()
    $local = $base
    $i = 1

    while ($true) {
        $mail = "$local@$DomainSuffix"
        $smtp1 = "smtp:$mail"
        $smtp2 = "SMTP:$mail"

        $exists = Get-ADObject -LDAPFilter "(|(mail=$mail)(proxyAddresses=$smtp1)(proxyAddresses=$smtp2))" -SearchBase $searchBase -ErrorAction SilentlyContinue
        if (-not $exists) { return $local }

        $i++
        $local = $base + $i
    }
}
