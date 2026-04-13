function Register-MainFormEvents {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Ui,
        [Parameter(Mandatory = $true)][hashtable]$Texts
    )

    Set-LogTarget -TextBox $Ui.TxtLog

    $script:LoadedUsers = $null
    $script:LoadedSheet = $null

    try { $Ui.TxtDomain.Text = (Get-ADDomain).DNSRoot } catch {}
    $Ui.TxtPdf.Text = Get-DefaultPasswordPdfPath -DomainSuffix $Ui.TxtDomain.Text

    $Ui.BtnExcel.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = Convert-UiText $Texts.FilterExcel
        if ($dlg.ShowDialog() -ne "OK") { return }

        $Ui.TxtExcel.Text = $dlg.FileName
        $script:LoadedUsers = $null
        $script:LoadedSheet = $null

        try {
            Write-Log ("{0}{1}" -f (Convert-UiText $Texts.LogOpenExcel), $Ui.TxtExcel.Text) "INFO"
            $res = Import-UsersFromExcelSmart -Path $Ui.TxtExcel.Text
            $script:LoadedUsers = $res.Users
            $script:LoadedSheet = $res.Sheet

            Write-Log ("{0}{1}. {2}{3}" -f (Convert-UiText $Texts.LogSheet), $script:LoadedSheet, (Convert-UiText $Texts.LogRows), $script:LoadedUsers.Count) "OK"
        }
        catch {
            Write-Log ("{0}{1}" -f (Convert-UiText $Texts.LogExcelError), $_.Exception.Message) "ERROR"
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, (Convert-UiText $Texts.CapExcelError), "OK", "Error") | Out-Null
        }
    })

    $Ui.BtnPdf.Add_Click({
        $saveDlg = New-Object System.Windows.Forms.SaveFileDialog
        $saveDlg.Filter = Convert-UiText $Texts.FilterPdf
        $saveDlg.FileName = [System.IO.Path]::GetFileName($Ui.TxtPdf.Text)
        $currentDir = [System.IO.Path]::GetDirectoryName($Ui.TxtPdf.Text)
        if (-not [string]::IsNullOrWhiteSpace($currentDir) -and (Test-Path -LiteralPath $currentDir)) {
            $saveDlg.InitialDirectory = $currentDir
        }
        if ($saveDlg.ShowDialog() -eq "OK") {
            $Ui.TxtPdf.Text = $saveDlg.FileName
            Write-Log ("{0}{1}" -f (Convert-UiText $Texts.LogPdfPath), $Ui.TxtPdf.Text) "OK"
        }
    })

    $Ui.BtnOU.Add_Click({
        try {
            $ou = Select-OU
            if ($ou) {
                $Ui.TxtOU.Text = $ou
                Write-Log ("{0}{1}" -f (Convert-UiText $Texts.LogOu), $ou) "OK"
            }
        }
        catch {
            Write-Log ("{0}{1}" -f (Convert-UiText $Texts.LogOuError), $_.Exception.Message) "ERROR"
        }
    })

    $Ui.BtnGroups.Add_Click({
        try {
            $sel = Select-Groups
            if ($sel -and $sel.Count -gt 0) {
                $Ui.TxtGroups.Text = ($sel -join ",")
                Write-Log ("{0}{1}" -f (Convert-UiText $Texts.LogGroups), ($sel -join ', ')) "OK"
            }
        }
        catch {
            Write-Log ("{0}{1}" -f (Convert-UiText $Texts.LogGroupsErr), $_.Exception.Message) "ERROR"
        }
    })

    $Ui.BtnRun.Add_Click({
        if (-not $script:LoadedUsers) {
            [System.Windows.Forms.MessageBox]::Show((Convert-UiText $Texts.MsgNoData), (Convert-UiText $Texts.CapNoData), "OK", "Warning") | Out-Null
            return
        }
        if ([string]::IsNullOrWhiteSpace($Ui.TxtOU.Text)) {
            [System.Windows.Forms.MessageBox]::Show((Convert-UiText $Texts.MsgNoOu), (Convert-UiText $Texts.CapNoOu), "OK", "Warning") | Out-Null
            return
        }
        if ([string]::IsNullOrWhiteSpace($Ui.TxtDomain.Text)) {
            [System.Windows.Forms.MessageBox]::Show((Convert-UiText $Texts.MsgNoDomain), (Convert-UiText $Texts.CapNoDomain), "OK", "Warning") | Out-Null
            return
        }
        if ([string]::IsNullOrWhiteSpace($Ui.TxtPdf.Text)) {
            [System.Windows.Forms.MessageBox]::Show((Convert-UiText $Texts.MsgNoPdf), (Convert-UiText $Texts.CapNoPdf), "OK", "Warning") | Out-Null
            return
        }

        $groups = @()
        if (-not [string]::IsNullOrWhiteSpace($Ui.TxtGroups.Text)) {
            $groups = $Ui.TxtGroups.Text.Split(',', [System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.Trim() }
        }

        Write-Log ("{0} OU={1}, {2}{3}, {4}{5}" -f (Convert-UiText $Texts.LogStart), $Ui.TxtOU.Text, (Convert-UiText $Texts.LogDomain), $Ui.TxtDomain.Text, (Convert-UiText $Texts.LogGroups2), ($groups -join ', ')) "INFO"

        $result = Create-UsersFromExcelData `
            -Users $script:LoadedUsers `
            -OU $Ui.TxtOU.Text `
            -DomainSuffix $Ui.TxtDomain.Text.Trim() `
            -GroupsToAdd $groups `
            -PasswordNeverExpires ([bool]$Ui.ChkNever.Checked)

        $createdRows = @($result | Where-Object { $_.Status -eq 'OK' -and -not [string]::IsNullOrWhiteSpace($_.Password) })
        if ($createdRows.Count -gt 0) {
            try {
                $pdfPath = Save-PasswordCredentialsPdf `
                    -Rows $createdRows `
                    -DomainSuffix $Ui.TxtDomain.Text.Trim() `
                    -OU $Ui.TxtOU.Text `
                    -PdfPath $Ui.TxtPdf.Text

                if ($pdfPath) {
                    Write-Log ("{0}{1}" -f (Convert-UiText $Texts.LogPdfSaved), $pdfPath) "OK"
                    $openPdf = [System.Windows.Forms.MessageBox]::Show(
                        ((Convert-UiText $Texts.MsgPdfSaved) + $pdfPath + (Convert-UiText $Texts.MsgOpenNow)),
                        (Convert-UiText $Texts.CapPdfSaved),
                        [System.Windows.Forms.MessageBoxButtons]::YesNo,
                        [System.Windows.Forms.MessageBoxIcon]::Information
                    )
                    if ($openPdf -eq [System.Windows.Forms.DialogResult]::Yes) {
                        Start-Process -FilePath $pdfPath | Out-Null
                    }
                }
                else {
                    Write-Log (Convert-UiText $Texts.LogPdfCancel) "INFO"
                }
            }
            catch {
                Write-Log ("{0}{1}" -f (Convert-UiText $Texts.LogPdfError), $_.Exception.Message) "ERROR"
                [System.Windows.Forms.MessageBox]::Show(
                    $_.Exception.Message,
                    (Convert-UiText $Texts.CapPdfError),
                    "OK",
                    "Error"
                ) | Out-Null
            }
        }
    })
}
