function New-MainFormLayout {
    param([Parameter(Mandatory)][hashtable]$Texts)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Texts.WindowTitle
    $form.Size = New-Object System.Drawing.Size(1040, 780)
    $form.MinimumSize = New-Object System.Drawing.Size(1000, 740)
    $form.StartPosition = 'CenterScreen'
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 252)
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 10)

    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.AutoPopDelay = 20000
    $toolTip.InitialDelay = 500
    $toolTip.ReshowDelay = 200
    $toolTip.ShowAlways = $true

    $contentPanel = New-Object System.Windows.Forms.Panel
    $contentPanel.Location = New-Object System.Drawing.Point(18, 18)
    $contentPanel.Size = New-Object System.Drawing.Size(988, 706)
    $contentPanel.Anchor = 'Top,Bottom,Left,Right'
    $form.Controls.Add($contentPanel)

    $cardImport = New-Object System.Windows.Forms.GroupBox
    $cardImport.Text = $Texts.ImportTitle
    $cardImport.Location = New-Object System.Drawing.Point(0, 0)
    $cardImport.Size = New-Object System.Drawing.Size(988, 114)
    $cardImport.Anchor = 'Top,Left,Right'
    $cardImport.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $contentPanel.Controls.Add($cardImport)

    $lblExcel = New-Object System.Windows.Forms.Label
    $lblExcel.Text = $Texts.ImportHint
    $lblExcel.Location = New-Object System.Drawing.Point(18, 26)
    $lblExcel.AutoSize = $true
    $lblExcel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $cardImport.Controls.Add($lblExcel)

    $txtExcel = New-Object System.Windows.Forms.TextBox
    $txtExcel.Location = New-Object System.Drawing.Point(18, 50)
    $txtExcel.Size = New-Object System.Drawing.Size(760, 25)
    $txtExcel.Anchor = 'Top,Left,Right'
    $cardImport.Controls.Add($txtExcel)

    $btnExcel = New-Object System.Windows.Forms.Button
    $btnExcel.Text = $Texts.ExcelButton
    $btnExcel.Location = New-Object System.Drawing.Point(806, 47)
    $btnExcel.Size = New-Object System.Drawing.Size(160, 32)
    $btnExcel.Anchor = 'Top,Right'
    $btnExcel.BackColor = [System.Drawing.Color]::FromArgb(228, 238, 255)
    $btnExcel.FlatStyle = 'Flat'
    $btnExcel.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(140, 170, 220)
    $cardImport.Controls.Add($btnExcel)

    $lblPdfFolder = New-Object System.Windows.Forms.Label
    $lblPdfFolder.Text = $Texts.PdfFolderLabel
    $lblPdfFolder.Location = New-Object System.Drawing.Point(18, 84)
    $lblPdfFolder.AutoSize = $true
    $lblPdfFolder.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $cardImport.Controls.Add($lblPdfFolder)

    $txtPdfFolder = New-Object System.Windows.Forms.TextBox
    $txtPdfFolder.Location = New-Object System.Drawing.Point(210, 80)
    $txtPdfFolder.Size = New-Object System.Drawing.Size(550, 25)
    $txtPdfFolder.Anchor = 'Top,Left,Right'
    $txtPdfFolder.Text = Get-DefaultPasswordPdfDirectory
    $cardImport.Controls.Add($txtPdfFolder)

    $btnPdfFolder = New-Object System.Windows.Forms.Button
    $btnPdfFolder.Text = $Texts.PdfFolderButton
    $btnPdfFolder.Location = New-Object System.Drawing.Point(806, 77)
    $btnPdfFolder.Size = New-Object System.Drawing.Size(160, 32)
    $btnPdfFolder.Anchor = 'Top,Right'
    $btnPdfFolder.FlatStyle = 'Flat'
    $btnPdfFolder.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(140, 170, 220)
    $cardImport.Controls.Add($btnPdfFolder)

    $cardSettings = New-Object System.Windows.Forms.GroupBox
    $cardSettings.Text = $Texts.SettingsTitle
    $cardSettings.Location = New-Object System.Drawing.Point(0, 126)
    $cardSettings.Size = New-Object System.Drawing.Size(988, 208)
    $cardSettings.Anchor = 'Top,Left,Right'
    $cardSettings.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $contentPanel.Controls.Add($cardSettings)

    $lblDom = New-Object System.Windows.Forms.Label
    $lblDom.Text = $Texts.DomainLabel
    $lblDom.Location = New-Object System.Drawing.Point(18, 34)
    $lblDom.AutoSize = $true
    $lblDom.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $cardSettings.Controls.Add($lblDom)

    $txtDomain = New-Object System.Windows.Forms.TextBox
    $txtDomain.Location = New-Object System.Drawing.Point(18, 58)
    $txtDomain.Size = New-Object System.Drawing.Size(285, 25)
    $cardSettings.Controls.Add($txtDomain)
    $toolTip.SetToolTip($txtDomain, $Texts.DomainTooltip)
    try { $txtDomain.Text = (Get-ADDomain).DNSRoot } catch {}

    $chkNever = New-Object System.Windows.Forms.CheckBox
    $chkNever.Text = $Texts.PasswordNeverExpires
    $chkNever.Location = New-Object System.Drawing.Point(330, 60)
    $chkNever.AutoSize = $true
    $chkNever.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $cardSettings.Controls.Add($chkNever)

    $lblOU = New-Object System.Windows.Forms.Label
    $lblOU.Text = $Texts.OuLabel
    $lblOU.Location = New-Object System.Drawing.Point(18, 98)
    $lblOU.AutoSize = $true
    $lblOU.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $cardSettings.Controls.Add($lblOU)

    $txtOU = New-Object System.Windows.Forms.TextBox
    $txtOU.Location = New-Object System.Drawing.Point(18, 122)
    $txtOU.Size = New-Object System.Drawing.Size(790, 25)
    $txtOU.Anchor = 'Top,Left,Right'
    $cardSettings.Controls.Add($txtOU)

    $btnOU = New-Object System.Windows.Forms.Button
    $btnOU.Text = $Texts.OuButton
    $btnOU.Location = New-Object System.Drawing.Point(842, 116)
    $btnOU.Size = New-Object System.Drawing.Size(124, 36)
    $btnOU.Anchor = 'Top,Right'
    $btnOU.FlatStyle = 'Flat'
    $btnOU.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(140, 170, 220)
    $cardSettings.Controls.Add($btnOU)

    $lblGroups = New-Object System.Windows.Forms.Label
    $lblGroups.Text = $Texts.GroupsLabel
    $lblGroups.Location = New-Object System.Drawing.Point(18, 148)
    $lblGroups.AutoSize = $true
    $lblGroups.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $cardSettings.Controls.Add($lblGroups)

    $txtGroups = New-Object System.Windows.Forms.TextBox
    $txtGroups.Location = New-Object System.Drawing.Point(18, 172)
    $txtGroups.Size = New-Object System.Drawing.Size(790, 25)
    $txtGroups.Anchor = 'Top,Left,Right'
    $cardSettings.Controls.Add($txtGroups)
    $toolTip.SetToolTip($txtGroups, $Texts.GroupsTooltip)

    $btnGroups = New-Object System.Windows.Forms.Button
    $btnGroups.Text = $Texts.GroupsButton
    $btnGroups.Location = New-Object System.Drawing.Point(842, 166)
    $btnGroups.Size = New-Object System.Drawing.Size(124, 36)
    $btnGroups.Anchor = 'Top,Right'
    $btnGroups.FlatStyle = 'Flat'
    $btnGroups.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(140, 170, 220)
    $cardSettings.Controls.Add($btnGroups)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = $Texts.RunButton
    $btnRun.Location = New-Object System.Drawing.Point(0, 358)
    $btnRun.Size = New-Object System.Drawing.Size(988, 52)
    $btnRun.Anchor = 'Top,Left,Right'
    $btnRun.BackColor = [System.Drawing.Color]::FromArgb(38, 125, 84)
    $btnRun.ForeColor = [System.Drawing.Color]::White
    $btnRun.FlatStyle = 'Flat'
    $btnRun.FlatAppearance.BorderSize = 0
    $btnRun.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
    $contentPanel.Controls.Add($btnRun)

    $cardLog = New-Object System.Windows.Forms.GroupBox
    $cardLog.Text = $Texts.LogTitle
    $cardLog.Location = New-Object System.Drawing.Point(0, 422)
    $cardLog.Size = New-Object System.Drawing.Size(988, 284)
    $cardLog.Anchor = 'Top,Bottom,Left,Right'
    $cardLog.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $contentPanel.Controls.Add($cardLog)

    $txtLog = New-Object System.Windows.Forms.TextBox
    $txtLog.Location = New-Object System.Drawing.Point(16, 28)
    $txtLog.Size = New-Object System.Drawing.Size(956, 246)
    $txtLog.Anchor = 'Top,Bottom,Left,Right'
    $txtLog.Multiline = $true
    $txtLog.ReadOnly = $true
    $txtLog.ScrollBars = 'Vertical'
    $txtLog.BackColor = [System.Drawing.Color]::White
    $txtLog.Font = New-Object System.Drawing.Font('Consolas', 9)
    $cardLog.Controls.Add($txtLog)

    $contentPanel.Add_SizeChanged({
        $cardLog.Height = [Math]::Max(160, $contentPanel.Height - 422)
        $txtLog.Width = [Math]::Max(200, $cardLog.ClientSize.Width - 32)
        $txtLog.Height = [Math]::Max(120, $cardLog.ClientSize.Height - 44)
    })

    Set-LogTarget -TextBox $txtLog

    return @{
        Form         = $form
        TxtExcel     = $txtExcel
        BtnExcel     = $btnExcel
        TxtPdfFolder = $txtPdfFolder
        BtnPdfFolder = $btnPdfFolder
        TxtDomain    = $txtDomain
        ChkNever     = $chkNever
        TxtOU        = $txtOU
        BtnOU        = $btnOU
        TxtGroups    = $txtGroups
        BtnGroups    = $btnGroups
        BtnRun       = $btnRun
        TxtLog       = $txtLog
    }
}
