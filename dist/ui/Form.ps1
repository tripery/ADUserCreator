function Show-MainForm {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Створення користувачів AD (Excel → AD)"
    $form.Size = New-Object System.Drawing.Size(980,780)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font("Segoe UI",10)

    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.AutoPopDelay = 20000
    $toolTip.InitialDelay = 500
    $toolTip.ReshowDelay = 200
    $toolTip.ShowAlways = $true

    # --- Excel block ---
    $lblExcel = New-Object System.Windows.Forms.Label
    $lblExcel.Text = "1) Excel (*.xlsx) з колонками: Вступник, Структурний підрозділ"
    $lblExcel.Location = New-Object System.Drawing.Point(20,18)
    $lblExcel.AutoSize = $true
    $form.Controls.Add($lblExcel)

    $txtExcel = New-Object System.Windows.Forms.TextBox
    $txtExcel.Location = New-Object System.Drawing.Point(20,45)
    $txtExcel.Size = New-Object System.Drawing.Size(700,25)
    $form.Controls.Add($txtExcel)

    $btnExcel = New-Object System.Windows.Forms.Button
    $btnExcel.Text = "Вибрати..."
    $btnExcel.Location = New-Object System.Drawing.Point(740,43)
    $btnExcel.Size = New-Object System.Drawing.Size(200,32)
    $form.Controls.Add($btnExcel)

    # Preview grid
    $lblPrev = New-Object System.Windows.Forms.Label
    $lblPrev.Text = "Preview (перші рядки з правильного листа):"
    $lblPrev.Location = New-Object System.Drawing.Point(20,82)
    $lblPrev.AutoSize = $true
    $form.Controls.Add($lblPrev)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(20,108)
    $grid.Size = New-Object System.Drawing.Size(920,220)
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AutoSizeColumnsMode = "Fill"
    $form.Controls.Add($grid)

    # --- Settings group ---
    $gb = New-Object System.Windows.Forms.GroupBox
    $gb.Text = "2) Налаштування AD"
    $gb.Location = New-Object System.Drawing.Point(20,340)
    $gb.Size = New-Object System.Drawing.Size(920,170)
    $form.Controls.Add($gb)

    # Domain
    $lblDom = New-Object System.Windows.Forms.Label
    $lblDom.Text = "Домен для UPN/пошти:"
    $lblDom.Location = New-Object System.Drawing.Point(15,30)
    $lblDom.AutoSize = $true
    $gb.Controls.Add($lblDom)

    $txtDomain = New-Object System.Windows.Forms.TextBox
    $txtDomain.Location = New-Object System.Drawing.Point(210,28)
    $txtDomain.Size = New-Object System.Drawing.Size(260,25)
    $gb.Controls.Add($txtDomain)
    $toolTip.SetToolTip($txtDomain, "Напр.: donnu.edu.ua")

    try { $txtDomain.Text = (Get-ADDomain).DNSRoot } catch {}

    # Never expire
    $chkNever = New-Object System.Windows.Forms.CheckBox
    $chkNever.Text = "Пароль ніколи не закінчується"
    $chkNever.Location = New-Object System.Drawing.Point(500,30)
    $gb.Controls.Add($chkNever)

    # OU
    $lblOU = New-Object System.Windows.Forms.Label
    $lblOU.Text = "OU для розміщення:"
    $lblOU.Location = New-Object System.Drawing.Point(15,68)
    $lblOU.AutoSize = $true
    $gb.Controls.Add($lblOU)

    $txtOU = New-Object System.Windows.Forms.TextBox
    $txtOU.Location = New-Object System.Drawing.Point(210,66)
    $txtOU.Size = New-Object System.Drawing.Size(520,25)
    $gb.Controls.Add($txtOU)

    $btnOU = New-Object System.Windows.Forms.Button
    $btnOU.Text = "Вибрати OU..."
    $btnOU.Location = New-Object System.Drawing.Point(745,64)
    $btnOU.Size = New-Object System.Drawing.Size(160,32)
    $gb.Controls.Add($btnOU)

    # Groups
    $lblGroups = New-Object System.Windows.Forms.Label
    $lblGroups.Text = "Групи (SamAccountName через кому):"
    $lblGroups.Location = New-Object System.Drawing.Point(15,106)
    $lblGroups.AutoSize = $true
    $gb.Controls.Add($lblGroups)

    $txtGroups = New-Object System.Windows.Forms.TextBox
    $txtGroups.Location = New-Object System.Drawing.Point(15,132)
    $txtGroups.Size = New-Object System.Drawing.Size(715,25)
    $gb.Controls.Add($txtGroups)
    $toolTip.SetToolTip($txtGroups, "Введи групи через кому або натисни 'Вибрати групи...'")

    $btnGroups = New-Object System.Windows.Forms.Button
    $btnGroups.Text = "Вибрати групи..."
    $btnGroups.Location = New-Object System.Drawing.Point(745,130)
    $btnGroups.Size = New-Object System.Drawing.Size(160,32)
    $gb.Controls.Add($btnGroups)

    # Run
    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = "СТВОРИТИ КОРИСТУВАЧІВ"
    $btnRun.Location = New-Object System.Drawing.Point(20,520)
    $btnRun.Size = New-Object System.Drawing.Size(920,50)
    $btnRun.BackColor = [System.Drawing.Color]::FromArgb(204,255,204)
    $btnRun.Font = New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($btnRun)

    # Log
    $lblLog = New-Object System.Windows.Forms.Label
    $lblLog.Text = "Журнал виконання:"
    $lblLog.Location = New-Object System.Drawing.Point(20,585)
    $lblLog.AutoSize = $true
    $form.Controls.Add($lblLog)

    $txtLog = New-Object System.Windows.Forms.TextBox
    $txtLog.Location = New-Object System.Drawing.Point(20,610)
    $txtLog.Size = New-Object System.Drawing.Size(920,120)
    $txtLog.Multiline = $true
    $txtLog.ReadOnly = $true
    $txtLog.ScrollBars = "Vertical"
    $txtLog.Font = New-Object System.Drawing.Font("Segoe UI",10)
    $form.Controls.Add($txtLog)

    Set-LogTarget -TextBox $txtLog

    # ---------- State ----------
    $script:LoadedUsers = $null
    $script:LoadedSheet = $null

    # ---------- Events ----------
    $btnExcel.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = "Excel файли (*.xlsx)|*.xlsx|Усі файли (*.*)|*.*"
        if ($dlg.ShowDialog() -ne "OK") { return }

        $txtExcel.Text = $dlg.FileName
        $grid.DataSource = $null
        $script:LoadedUsers = $null
        $script:LoadedSheet = $null

        try {
            Write-Log "Відкриття Excel: $($txtExcel.Text)" "INFO"
            $res = Import-UsersFromExcelSmart -Path $txtExcel.Text
            $script:LoadedUsers = $res.Users
            $script:LoadedSheet = $res.Sheet

            Write-Log "Використовується лист: $($script:LoadedSheet). Рядків: $($script:LoadedUsers.Count)" "OK"

            # preview top 20
            $preview = $script:LoadedUsers | Select-Object -First 20
            $grid.DataSource = $preview
        }
        catch {
            Write-Log "Помилка Excel: $($_.Exception.Message)" "ERROR"
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Excel помилка", "OK", "Error") | Out-Null
        }
    })

    $btnOU.Add_Click({
        try {
            $ou = Select-OU
            if ($ou) { $txtOU.Text = $ou; Write-Log "OU вибрано: $ou" "OK" }
        } catch {
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
        } catch {
            Write-Log "Помилка вибору груп: $($_.Exception.Message)" "ERROR"
        }
    })

    $btnRun.Add_Click({
        if (-not $script:LoadedUsers) { [System.Windows.Forms.MessageBox]::Show("Спочатку вибери Excel файл.", "Немає даних", "OK", "Warning") | Out-Null; return }
        if ([string]::IsNullOrWhiteSpace($txtOU.Text)) { [System.Windows.Forms.MessageBox]::Show("Вибери OU.", "Немає OU", "OK", "Warning") | Out-Null; return }
        if ([string]::IsNullOrWhiteSpace($txtDomain.Text)) { [System.Windows.Forms.MessageBox]::Show("Вкажи домен для UPN/пошти.", "Немає домену", "OK", "Warning") | Out-Null; return }

        $groups = @()
        if (-not [string]::IsNullOrWhiteSpace($txtGroups.Text)) {
            $groups = $txtGroups.Text.Split(',', [System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.Trim() }
        }

        Write-Log "СТАРТ: створення користувачів. OU=$($txtOU.Text), Domain=$($txtDomain.Text), Groups=$($groups -join ', ')" "INFO"

        Create-UsersFromExcelData `
            -Users $script:LoadedUsers `
            -OU $txtOU.Text `
            -DomainSuffix $txtDomain.Text.Trim() `
            -GroupsToAdd $groups `
            -PasswordNeverExpires ([bool]$chkNever.Checked)
    })

    $null = $form.ShowDialog()
}
