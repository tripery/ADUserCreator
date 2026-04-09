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
    }
    if ($Preview.middleName) { $newParams['OtherName'] = $Preview.middleName }
    if ($PasswordNeverExpires) {
        $newParams['PasswordNeverExpires'] = $true
        $newParams['ChangePasswordAtLogon'] = $false
    } else {
        $newParams['ChangePasswordAtLogon'] = $true
    }
    New-ADUser @newParams
    foreach ($groupSam in $GroupsToAdd) { if ($groupSam) { Add-ADGroupMember -Identity $groupSam -Members $Preview.login -ErrorAction Stop } }
    [pscustomobject]@{ fullName = $Preview.fullName; login = $Preview.login; email = $Preview.email; password = $password; unit = $Preview.unit; status = 'created' }
}
