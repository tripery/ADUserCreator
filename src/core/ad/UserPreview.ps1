function Split-FullName {
    param([Parameter(Mandatory)][string]$FullName)
    $parts = (Normalize-Text $FullName) -split '\s+' | Where-Object { $_ }
    if ($parts.Count -lt 2) { throw ("Unable to parse fullName '{0}'. Expected at least surname and given name." -f $FullName) }
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
    if ([string]::IsNullOrWhiteSpace($surnameLat) -or [string]::IsNullOrWhiteSpace($givenLat)) { throw 'Failed to transliterate full name to Latin.' }
    $samBase = ('{0}.{1}' -f $givenLat.Substring(0,1), $surnameLat).ToLower()
    [pscustomobject]@{ SurnameLatin = $surnameLat; GivenLatin = $givenLat; MiddleLatin = $middleLat; SamBase = $samBase; MailLocalBase = $samBase }
}

function New-PreviewUserRecord {
    param([Parameter(Mandatory)]$UserItem,[Parameter(Mandatory)][string]$DomainSuffix,[string]$OU,[switch]$CheckUniqueness)
    $fullName = Normalize-Text $UserItem.fullName
    if ([string]::IsNullOrWhiteSpace($fullName)) { throw 'Empty fullName field in user record.' }
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
