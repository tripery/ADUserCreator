function Resolve-InputFullName {
    param([Parameter(Mandatory=$true)]$User)

    $candidates = @(
        $User.'Вступник',
        $User.'ПIБ',
        $User.'ПІБ',
        $User.FullName,
        $User.Name
    )

    foreach ($value in $candidates) {
        $s = [string]$value
        if (-not [string]::IsNullOrWhiteSpace($s)) {
            return (($s -replace '\s+', ' ').Trim())
        }
    }

    return $null
}

function Get-RowValue {
    param(
        [Parameter(Mandatory=$true)]$User,
        [Parameter(Mandatory=$true)][string[]]$Names
    )

    foreach ($name in $Names) {
        $prop = $User.PSObject.Properties[$name]
        if ($null -eq $prop) { continue }

        $value = [string]$prop.Value
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value.Trim()
        }
    }

    return $null
}

function Create-UsersFromExcelData {
    param(
        [Parameter(Mandatory=$true)]$Users,
        [Parameter(Mandatory=$true)][string]$OU,
        [Parameter(Mandatory=$true)][string]$DomainSuffix,
        [string[]]$GroupsToAdd = @(),
        [bool]$PasswordNeverExpires = $false
    )

    $domain = $DomainSuffix.Trim().ToLower()
    $groupList = @($GroupsToAdd | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() } | Select-Object -Unique)

    $total = @($Users).Count
    $okCount = 0
    $errCount = 0
    $result = @()

    Write-Log "Початок обробки: $total користувач(ів)." "INFO"

    $rowIndex = 0
    foreach ($row in $Users) {
        $rowIndex++

        try {
            $displayName = Resolve-InputFullName -User $row
            if ([string]::IsNullOrWhiteSpace($displayName)) {
                throw "Рядок ${rowIndex}: не заповнено ПІБ (колонка 'Вступник')."
            }

            $parts = $displayName.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)
            if ($parts.Count -lt 2) {
                throw "Рядок ${rowIndex}: очікується щонайменше 2 частини ПІБ, отримано '$displayName'."
            }

            $surname = $parts[0]
            $givenName = $parts[1]

            $surnameLat = (Convert-UA2Latin $surname) -replace '[^a-z0-9]', ''
            $givenLat = (Convert-UA2Latin $givenName) -replace '[^a-z0-9]', ''

            if ([string]::IsNullOrWhiteSpace($surnameLat)) { $surnameLat = "user$rowIndex" }
            if ([string]::IsNullOrWhiteSpace($givenLat)) { $givenLat = "u" }

            $baseSam = "$surnameLat.$givenLat"
            $sam = Get-UniqueSamAccountName -BaseSam $baseSam

            $givenInitial = $givenLat.Substring(0, 1)
            $mailBase = "$surnameLat.$givenInitial"
            $mailLocal = Get-UniqueMailLocalPart -BaseLocal $mailBase -DomainSuffix $domain
            $mail = "$mailLocal@$domain"
            $upn = "$sam@$domain"

            $cn = Get-UniqueCN -BaseCN $displayName -OUPath $OU

            $plainPassword = Get-RandomPassword
            $securePassword = ConvertTo-SecureString $plainPassword -AsPlainText -Force

            $title = Get-RowValue -User $row -Names @('Посада','Должность','Title','Job Title')
            $department = Get-RowValue -User $row -Names @('Відділ','Отдел','Структурний підрозділ','Підрозділ','Department')
            $company = Get-RowValue -User $row -Names @('Організація','Организация','Company')

            $newUserParams = @{
                Name                  = $cn
                DisplayName           = $displayName
                GivenName             = $givenName
                Surname               = $surname
                SamAccountName        = $sam
                UserPrincipalName     = $upn
                EmailAddress          = $mail
                Path                  = $OU
                Enabled               = $true
                AccountPassword       = $securePassword
                ChangePasswordAtLogon = (-not [bool]$PasswordNeverExpires)
                PasswordNeverExpires  = $PasswordNeverExpires
                ErrorAction           = 'Stop'
            }

            if (-not [string]::IsNullOrWhiteSpace($title)) { $newUserParams['Title'] = $title }
            if (-not [string]::IsNullOrWhiteSpace($department)) { $newUserParams['Department'] = $department }
            if (-not [string]::IsNullOrWhiteSpace($company)) { $newUserParams['Company'] = $company }

            New-ADUser @newUserParams

            # Set SMTP + SIP proxy addresses right after user creation.
            $proxyAddresses = @("SMTP:$mail", "sip:$mail")
            Set-ADUser -Identity $sam -Add @{ proxyAddresses = $proxyAddresses } -ErrorAction Stop

            foreach ($groupName in $groupList) {
                try {
                    Add-ADGroupMember -Identity $groupName -Members $sam -ErrorAction Stop
                }
                catch {
                    Write-Log "[$displayName] Не вдалося додати в групу '$groupName': $($_.Exception.Message)" "WARN"
                }
            }

            $okCount++
            $result += [pscustomobject]@{
                Row               = $rowIndex
                DisplayName       = $displayName
                SamAccountName    = $sam
                UserPrincipalName = $upn
                Mail              = $mail
                Password          = $plainPassword
                Status            = 'OK'
            }

            Write-Log "[$displayName] Створено: Sam=$sam, UPN=$upn, Mail=$mail, Pass=$plainPassword" "OK"
        }
        catch {
            $errCount++
            $result += [pscustomobject]@{
                Row         = $rowIndex
                DisplayName = (Resolve-InputFullName -User $row)
                Status      = 'ERROR'
                Error       = $_.Exception.Message
            }

            Write-Log "Рядок ${rowIndex}: $($_.Exception.Message)" "ERROR"
        }
    }

    Write-Log "Завершено. Успішно: $okCount; Помилок: $errCount." "INFO"
    return $result
}
