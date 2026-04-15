function Register-MainFormEvents {
    param(
        [Parameter(Mandatory)][hashtable]$Ui,
        [Parameter(Mandatory)][hashtable]$Texts
    )

    $script:LoadedUsers = $null
    $script:LoadedSheet = $null

    $Ui.BtnExcel.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = 'Excel файли (*.xlsx)|*.xlsx|Усі файли (*.*)|*.*'
        if ($dlg.ShowDialog() -ne 'OK') { return }

        $Ui.TxtExcel.Text = $dlg.FileName
        $script:LoadedUsers = $null
        $script:LoadedSheet = $null

        try {
            Write-Log "Відкриття Excel: $($Ui.TxtExcel.Text)" 'INFO'
            $res = Import-UsersFromExcelSmart -Path $Ui.TxtExcel.Text
            $script:LoadedUsers = $res.Users
            $script:LoadedSheet = $res.Sheet
            Write-Log "Використовується лист: $($script:LoadedSheet). Рядків: $($script:LoadedUsers.Count)" 'OK'
        }
        catch {
            Write-Log "Помилка Excel: $($_.Exception.Message)" 'ERROR'
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, $Texts.ExcelErrorTitle, 'OK', 'Error') | Out-Null
        }
    })

    $Ui.BtnPdfFolder.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        if (-not [string]::IsNullOrWhiteSpace($Ui.TxtPdfFolder.Text) -and (Test-Path -LiteralPath $Ui.TxtPdfFolder.Text)) {
            $dlg.SelectedPath = $Ui.TxtPdfFolder.Text
        } else {
            $dlg.SelectedPath = Get-DefaultPasswordPdfDirectory
        }

        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $Ui.TxtPdfFolder.Text = $dlg.SelectedPath
            Write-Log "Папку для PDF вибрано: $($Ui.TxtPdfFolder.Text)" 'OK'
        }
    })

    $Ui.BtnOU.Add_Click({
        try {
            $ou = Select-OU
            if ($ou) {
                $Ui.TxtOU.Text = $ou
                Write-Log "OU вибрано: $ou" 'OK'
            }
        }
        catch {
            Write-Log "Помилка OU picker: $($_.Exception.Message)" 'ERROR'
        }
    })

    $Ui.BtnGroups.Add_Click({
        try {
            $sel = Select-Groups
            if ($sel -and $sel.Count -gt 0) {
                $Ui.TxtGroups.Text = ($sel -join ',')
                Write-Log "Групи вибрано: $($sel -join ', ')" 'OK'
            }
        }
        catch {
            Write-Log "Помилка вибору груп: $($_.Exception.Message)" 'ERROR'
        }
    })

    $Ui.BtnRun.Add_Click({
        if (-not $script:LoadedUsers) {
            [System.Windows.Forms.MessageBox]::Show('Спочатку виберіть Excel файл.', $Texts.NoDataTitle, 'OK', 'Warning') | Out-Null
            return
        }
        if ([string]::IsNullOrWhiteSpace($Ui.TxtOU.Text)) {
            [System.Windows.Forms.MessageBox]::Show('Виберіть OU.', $Texts.NoOuTitle, 'OK', 'Warning') | Out-Null
            return
        }
        if ([string]::IsNullOrWhiteSpace($Ui.TxtDomain.Text)) {
            [System.Windows.Forms.MessageBox]::Show('Вкажіть домен для UPN/пошти.', $Texts.NoDomainTitle, 'OK', 'Warning') | Out-Null
            return
        }
        if ([string]::IsNullOrWhiteSpace($Ui.TxtPdfFolder.Text)) {
            [System.Windows.Forms.MessageBox]::Show('Виберіть папку для зберігання PDF з паролями.', $Texts.NoPdfFolderTitle, 'OK', 'Warning') | Out-Null
            return
        }

        $groups = @()
        if (-not [string]::IsNullOrWhiteSpace($Ui.TxtGroups.Text)) {
            $groups = $Ui.TxtGroups.Text.Split(',', [System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.Trim() }
        }

        Write-Log "СТАРТ: створення користувачів. OU=$($Ui.TxtOU.Text), Domain=$($Ui.TxtDomain.Text), Groups=$($groups -join ', ')" 'INFO'

        $result = Create-UsersFromExcelData `
            -Users $script:LoadedUsers `
            -OU $Ui.TxtOU.Text `
            -DomainSuffix $Ui.TxtDomain.Text.Trim() `
            -GroupsToAdd $groups `
            -PasswordNeverExpires ([bool]$Ui.ChkNever.Checked)

        $createdRows = @($result | Where-Object { $_.Status -eq 'OK' -and -not [string]::IsNullOrWhiteSpace($_.Password) })
        if ($createdRows.Count -gt 0) {
            try {
                $pdfPath = Save-PasswordCredentialsPdfToFolder `
                    -Rows $createdRows `
                    -DomainSuffix $Ui.TxtDomain.Text.Trim() `
                    -OU $Ui.TxtOU.Text `
                    -OutputDirectory $Ui.TxtPdfFolder.Text

                if ($pdfPath) {
                    Write-Log "PDF з паролями збережено: $pdfPath" 'OK'
                }
            }
            catch {
                Write-Log "Помилка створення PDF: $($_.Exception.Message)" 'ERROR'
                [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, $Texts.PdfErrorTitle, 'OK', 'Error') | Out-Null
            }
        }
    })
}
