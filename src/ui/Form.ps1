function Show-MainForm {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "ADUserCreator"
    $form.Size = New-Object System.Drawing.Size(980, 720)
    $form.MinimumSize = New-Object System.Drawing.Size(980, 720)
    $form.StartPosition = "CenterScreen"
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 255)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.AutoPopDelay = 20000
    $toolTip.InitialDelay = 500
    $toolTip.ReshowDelay = 200
    $toolTip.ShowAlways = $true

    $panelHero = New-Object System.Windows.Forms.Panel
    $panelHero.Location = New-Object System.Drawing.Point(20, 18)
    $panelHero.Size = New-Object System.Drawing.Size(924, 102)
    $panelHero.BackColor = [System.Drawing.Color]::FromArgb(28, 72, 156)
    $form.Controls.Add($panelHero)

    $lblHeroKicker = New-Object System.Windows.Forms.Label
    $lblHeroKicker.Text = "ADUSERCREATOR"
    $lblHeroKicker.Location = New-Object System.Drawing.Point(24, 16)
    $lblHeroKicker.AutoSize = $true
    $lblHeroKicker.ForeColor = [System.Drawing.Color]::FromArgb(214, 230, 255)
    $lblHeroKicker.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $panelHero.Controls.Add($lblHeroKicker)

    $lblHeroTitle = New-Object System.Windows.Forms.Label
    $lblHeroTitle.Text = "Масове створення користувачів Active Directory"
    $lblHeroTitle.Location = New-Object System.Drawing.Point(22, 34)
    $lblHeroTitle.AutoSize = $true
    $lblHeroTitle.ForeColor = [System.Drawing.Color]::White
    $lblHeroTitle.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
    $panelHero.Controls.Add($lblHeroTitle)

    $lblHeroText = New-Object System.Windows.Forms.Label
    $lblHeroText.Text = "Завантажте Excel, оберіть OU, додайте групи та створіть облікові записи. Після успішного виконання можна одразу зберегти PDF з логінами і паролями."
    $lblHeroText.Location = New-Object System.Drawing.Point(26, 66)
    $lblHeroText.Size = New-Object System.Drawing.Size(760, 28)
    $lblHeroText.ForeColor = [System.Drawing.Color]::FromArgb(230, 238, 255)
    $panelHero.Controls.Add($lblHeroText)

    $panelStats = New-Object System.Windows.Forms.Panel
    $panelStats.Location = New-Object System.Drawing.Point(796, 16)
    $panelStats.Size = New-Object System.Drawing.Size(110, 68)
    $panelStats.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
    $panelHero.Controls.Add($panelStats)

    $lblStatsTitle = New-Object System.Windows.Forms.Label
    $lblStatsTitle.Text = "Рядків"
    $lblStatsTitle.Location = New-Object System.Drawing.Point(12, 10)
    $lblStatsTitle.AutoSize = $true
    $lblStatsTitle.ForeColor = [System.Drawing.Color]::FromArgb(88, 105, 140)
    $panelStats.Controls.Add($lblStatsTitle)

    $lblStatsValue = New-Object System.Windows.Forms.Label
    $lblStatsValue.Text = "0"
    $lblStatsValue.Location = New-Object System.Drawing.Point(12, 26)
    $lblStatsValue.AutoSize = $true
    $lblStatsValue.ForeColor = [System.Drawing.Color]::FromArgb(28, 72, 156)
    $lblStatsValue.Font = New-Object System.Drawing.Font("Segoe UI", 20, [System.Drawing.FontStyle]::Bold)
    $panelStats.Controls.Add($lblStatsValue)

    $lblExcel = New-Object System.Windows.Forms.Label
    $lblExcel.Text = "1) Excel (*.xlsx) з колонками: Вступник, Структурний підрозділ"
    $lblExcel.Location = New-Object System.Drawing.Point(20, 140)
    $lblExcel.AutoSize = $true
    $lblExcel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblExcel)

    $txtExcel = New-Object System.Windows.Forms.TextBox
    $txtExcel.Location = New-Object System.Drawing.Point(20, 168)
    $txtExcel.Size = New-Object System.Drawing.Size(700, 27)
    $txtExcel.BackColor = [System.Drawing.Color]::White
    $form.Controls.Add($txtExcel)

    $btnExcel = New-Object System.Windows.Forms.Button
    $btnExcel.Text = "Вибрати файл"
    $btnExcel.Location = New-Object System.Drawing.Point(740, 165)
    $btnExcel.Size = New-Object System.Drawing.Size(204, 34)
    $btnExcel.BackColor = [System.Drawing.Color]::FromArgb(234, 242, 255)
    $btnExcel.FlatStyle = 'Flat'
    $form.Controls.Add($btnExcel)

    $panelInfo = New-Object System.Windows.Forms.Panel
    $panelInfo.Location = New-Object System.Drawing.Point(20, 214)
    $panelInfo.Size = New-Object System.Drawing.Size(924, 72)
    $panelInfo.BackColor = [System.Drawing.Color]::White
    $form.Controls.Add($panelInfo)

    $lblInfoTitle = New-Object System.Windows.Forms.Label
    $lblInfoTitle.Text = "Поточний файл"
    $lblInfoTitle.Location = New-Object System.Drawing.Point(18, 12)
    $lblInfoTitle.AutoSize = $true
    $lblInfoTitle.ForeColor = [System.Drawing.Color]::FromArgb(98, 112, 138)
    $lblInfoTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $panelInfo.Controls.Add($lblInfoTitle)

    $lblFileStatus = New-Object System.Windows.Forms.Label
    $lblFileStatus.Text = "Файл не вибрано"
    $lblFileStatus.Location = New-Object System.Drawing.Point(18, 34)
    $lblFileStatus.Size = New-Object System.Drawing.Size(620, 24)
    $lblFileStatus.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
    $panelInfo.Controls.Add($lblFileStatus)

    $lblSheetStatus = New-Object System.Windows.Forms.Label
    $lblSheetStatus.Text = "Лист: —"
    $lblSheetStatus.Location = New-Object System.Drawing.Point(648, 18)
    $lblSheetStatus.Size = New-Object System.Drawing.Size(250, 20)
    $lblSheetStatus.ForeColor = [System.Drawing.Color]::FromArgb(98, 112, 138)
    $panelInfo.Controls.Add($lblSheetStatus)

    $lblRowsStatus = New-Object System.Windows.Forms.Label
    $lblRowsStatus.Text = "Записів: 0"
    $lblRowsStatus.Location = New-Object System.Drawing.Point(648, 42)
    $lblRowsStatus.Size = New-Object System.Drawing.Size(250, 20)
    $lblRowsStatus.ForeColor = [System.Drawing.Color]::FromArgb(98, 112, 138)
    $panelInfo.Controls.Add($lblRowsStatus)

    $gb = New-Object System.Windows.Forms.GroupBox
    $gb.Text = "2) Налаштування AD"
    $gb.Location = New-Object System.Drawing.Point(20, 304)
    $gb.Size = New-Object System.Drawing.Size(924, 190)
    $form.Controls.Add($gb)

    $lblDom = New-Object System.Windows.Forms.Label
    $lblDom.Text = "Домен для UPN / пошти"
    $lblDom.Location = New-Object System.Drawing.Point(18, 34)
    $lblDom.AutoSize = $true
    $gb.Controls.Add($lblDom)

    $txtDomain = New-Object System.Windows.Forms.TextBox
    $txtDomain.Location = New-Object System.Drawing.Point(18, 58)
    $txtDomain.Size = New-Object System.Drawing.Size(300, 27)
    $gb.Controls.Add($txtDomain)
    $toolTip.SetToolTip($txtDomain, "Наприклад: donnu.edu.ua")

    try { $txtDomain.Text = (Get-ADDomain).DNSRoot } catch {}

    $chkNever = New-Object System.Windows.Forms.CheckBox
    $chkNever.Text = "Пароль ніколи не закінчується"
    $chkNever.Location = New-Object System.Drawing.Point(350, 60)
    $chkNever.AutoSize = $true
    $gb.Controls.Add($chkNever)

    $lblOU = New-Object System.Windows.Forms.Label
    $lblOU.Text = "OU для розміщення"
    $lblOU.Location = New-Object System.Drawing.Point(18, 100)
    $lblOU.AutoSize = $true
    $gb.Controls.Add($lblOU)

    $txtOU = New-Object System.Windows.Forms.TextBox
    $txtOU.Location = New-Object System.Drawing.Point(18, 124)
    $txtOU.Size = New-Object System.Drawing.Size(720, 27)
    $gb.Controls.Add($txtOU)

    $btnOU = New-Object System.Windows.Forms.Button
    $btnOU.Text = "Вибрати OU"
    $btnOU.Location = New-Object System.Drawing.Point(756, 121)
    $btnOU.Size = New-Object System.Drawing.Size(148, 34)
    $btnOU.BackColor = [System.Drawing.Color]::FromArgb(243, 247, 255)
    $btnOU.FlatStyle = 'Flat'
    $gb.Controls.Add($btnOU)

    $lblGroups = New-Object System.Windows.Forms.Label
    $lblGroups.Text = "Групи (SamAccountName через кому)"
    $lblGroups.Location = New-Object System.Drawing.Point(18, 160)
    $lblGroups.AutoSize = $true
    $gb.Controls.Add($lblGroups)

    $txtGroups = New-Object System.Windows.Forms.TextBox
    $txtGroups.Location = New-Object System.Drawing.Point(320, 157)
    $txtGroups.Size = New-Object System.Drawing.Size(418, 27)
    $gb.Controls.Add($txtGroups)
    $toolTip.SetToolTip($txtGroups, "Введи групи через кому або натисни 'Вибрати групи'")

    $btnGroups = New-Object System.Windows.Forms.Button
    $btnGroups.Text = "Вибрати групи"
    $btnGroups.Location = New-Object System.Drawing.Point(756, 154)
    $btnGroups.Size = New-Object System.Drawing.Size(148, 34)
    $btnGroups.BackColor = [System.Drawing.Color]::FromArgb(243, 247, 255)
    $btnGroups.FlatStyle = 'Flat'
    $gb.Controls.Add($btnGroups)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = "СТВОРИТИ КОРИСТУВАЧІВ"
    $btnRun.Location = New-Object System.Drawing.Point(20, 514)
    $btnRun.Size = New-Object System.Drawing.Size(924, 50)
    $btnRun.BackColor = [System.Drawing.Color]::FromArgb(208, 245, 217)
    $btnRun.FlatStyle = 'Flat'
    $btnRun.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($btnRun)

    $lblLog = New-Object System.Windows.Forms.Label
    $lblLog.Text = "Журнал виконання"
    $lblLog.Location = New-Object System.Drawing.Point(20, 580)
    $lblLog.AutoSize = $true
    $lblLog.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblLog)

    $txtLog = New-Object System.Windows.Forms.TextBox
    $txtLog.Location = New-Object System.Drawing.Point(20, 606)
    $txtLog.Size = New-Object System.Drawing.Size(924, 70)
    $txtLog.Multiline = $true
    $txtLog.ReadOnly = $true
    $txtLog.ScrollBars = "Vertical"
    $txtLog.Font = New-Object System.Drawing.Font("Consolas", 9.5)
    $txtLog.BackColor = [System.Drawing.Color]::FromArgb(249, 251, 255)
    $form.Controls.Add($txtLog)

    Set-LogTarget -TextBox $txtLog

    $script:LoadedUsers = $null
    $script:LoadedSheet = $null

    $btnExcel.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = "Excel файли (*.xlsx)|*.xlsx|Усі файли (*.*)|*.*"
        if ($dlg.ShowDialog() -ne "OK") { return }

        $txtExcel.Text = $dlg.FileName
        $script:LoadedUsers = $null
        $script:LoadedSheet = $null
        $lblFileStatus.Text = "Завантаження Excel..."
        $lblSheetStatus.Text = "Лист: —"
        $lblRowsStatus.Text = "Записів: 0"
        $lblStatsValue.Text = "0"

        try {
            Write-Log "Відкриття Excel: $($txtExcel.Text)" "INFO"
            $res = Import-UsersFromExcelSmart -Path $txtExcel.Text
            $script:LoadedUsers = $res.Users
            $script:LoadedSheet = $res.Sheet

            $lblFileStatus.Text = [System.IO.Path]::GetFileName($txtExcel.Text)
            $lblSheetStatus.Text = "Лист: $($script:LoadedSheet)"
            $lblRowsStatus.Text = "Записів: $($script:LoadedUsers.Count)"
            $lblStatsValue.Text = [string]$script:LoadedUsers.Count

            Write-Log "Використовується лист: $($script:LoadedSheet). Рядків: $($script:LoadedUsers.Count)" "OK"
        }
        catch {
            $lblFileStatus.Text = "Помилка завантаження файлу"
            Write-Log "Помилка Excel: $($_.Exception.Message)" "ERROR"
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Excel помилка", "OK", "Error") | Out-Null
        }
    })

    $btnOU.Add_Click({
        try {
            $ou = Select-OU
            if ($ou) {
                $txtOU.Text = $ou
                Write-Log "OU вибрано: $ou" "OK"
            }
        }
        catch {
            Write-Log "Помилка OU picker: $($_.Exception.Message)" "ERROR"
        }
    })

    $btnGroups.Add_Click({
        try {
            $sel = Select-Groups
            if ($sel -and $sel.Count -gt 0) {
                $txtGroups.Text = ($sel -join ",")
                Write-Log "Групи вибрано: $($sel -join ', ')" "OK"
            }
        }
        catch {
            Write-Log "Помилка вибору груп: $($_.Exception.Message)" "ERROR"
        }
    })

    $btnRun.Add_Click({
        if (-not $script:LoadedUsers) {
            [System.Windows.Forms.MessageBox]::Show("Спочатку вибери Excel файл.", "Немає даних", "OK", "Warning") | Out-Null
            return
        }
        if ([string]::IsNullOrWhiteSpace($txtOU.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Вибери OU.", "Немає OU", "OK", "Warning") | Out-Null
            return
        }
        if ([string]::IsNullOrWhiteSpace($txtDomain.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Вкажи домен для UPN/пошти.", "Немає домену", "OK", "Warning") | Out-Null
            return
        }

        $groups = @()
        if (-not [string]::IsNullOrWhiteSpace($txtGroups.Text)) {
            $groups = $txtGroups.Text.Split(',', [System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.Trim() }
        }

        Write-Log "СТАРТ: створення користувачів. OU=$($txtOU.Text), Domain=$($txtDomain.Text), Groups=$($groups -join ', ')" "INFO"

        $result = Create-UsersFromExcelData `
            -Users $script:LoadedUsers `
            -OU $txtOU.Text `
            -DomainSuffix $txtDomain.Text.Trim() `
            -GroupsToAdd $groups `
            -PasswordNeverExpires ([bool]$chkNever.Checked)

        $createdRows = @($result | Where-Object { $_.Status -eq 'OK' -and -not [string]::IsNullOrWhiteSpace($_.Password) })
        if ($createdRows.Count -gt 0) {
            try {
                $pdfPath = Save-PasswordCredentialsPdfInteractive `
                    -Rows $createdRows `
                    -DomainSuffix $txtDomain.Text.Trim() `
                    -OU $txtOU.Text

                if ($pdfPath) {
                    Write-Log "PDF з паролями збережено: $pdfPath" "OK"
                    $openPdf = [System.Windows.Forms.MessageBox]::Show(
                        "PDF з паролями збережено.`r`n`r`n$pdfPath`r`n`r`nВідкрити файл зараз?",
                        "PDF збережено",
                        [System.Windows.Forms.MessageBoxButtons]::YesNo,
                        [System.Windows.Forms.MessageBoxIcon]::Information
                    )
                    if ($openPdf -eq [System.Windows.Forms.DialogResult]::Yes) {
                        Start-Process -FilePath $pdfPath | Out-Null
                    }
                }
                else {
                    Write-Log "Збереження PDF скасовано користувачем." "INFO"
                }
            }
            catch {
                Write-Log "Помилка створення PDF: $($_.Exception.Message)" "ERROR"
                [System.Windows.Forms.MessageBox]::Show(
                    $_.Exception.Message,
                    "PDF помилка",
                    "OK",
                    "Error"
                ) | Out-Null
            }
        }
    })

    $null = $form.ShowDialog()
}
