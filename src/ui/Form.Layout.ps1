function New-MainFormUi {
    param([Parameter(Mandatory = $true)][hashtable]$Texts)

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

    $lblExcel = New-Object System.Windows.Forms.Label
    $lblExcel.Text = Convert-UiText $Texts.LblExcel
    $lblExcel.Location = New-Object System.Drawing.Point(20, 24)
    $lblExcel.AutoSize = $true
    $lblExcel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblExcel)

    $txtExcel = New-Object System.Windows.Forms.TextBox
    $txtExcel.Location = New-Object System.Drawing.Point(20, 52)
    $txtExcel.Size = New-Object System.Drawing.Size(700, 27)
    $txtExcel.BackColor = [System.Drawing.Color]::White
    $form.Controls.Add($txtExcel)

    $btnExcel = New-Object System.Windows.Forms.Button
    $btnExcel.Text = Convert-UiText $Texts.BtnExcel
    $btnExcel.Location = New-Object System.Drawing.Point(740, 49)
    $btnExcel.Size = New-Object System.Drawing.Size(204, 34)
    $btnExcel.BackColor = [System.Drawing.Color]::FromArgb(234, 242, 255)
    $btnExcel.FlatStyle = 'Flat'
    $form.Controls.Add($btnExcel)

    $lblPdf = New-Object System.Windows.Forms.Label
    $lblPdf.Text = Convert-UiText $Texts.LblPdf
    $lblPdf.Location = New-Object System.Drawing.Point(20, 92)
    $lblPdf.AutoSize = $true
    $lblPdf.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblPdf)

    $txtPdf = New-Object System.Windows.Forms.TextBox
    $txtPdf.Location = New-Object System.Drawing.Point(20, 120)
    $txtPdf.Size = New-Object System.Drawing.Size(700, 27)
    $txtPdf.BackColor = [System.Drawing.Color]::White
    $form.Controls.Add($txtPdf)

    $btnPdf = New-Object System.Windows.Forms.Button
    $btnPdf.Text = Convert-UiText $Texts.BtnPdf
    $btnPdf.Location = New-Object System.Drawing.Point(740, 117)
    $btnPdf.Size = New-Object System.Drawing.Size(204, 34)
    $btnPdf.BackColor = [System.Drawing.Color]::FromArgb(243, 247, 255)
    $btnPdf.FlatStyle = 'Flat'
    $form.Controls.Add($btnPdf)

    $gb = New-Object System.Windows.Forms.GroupBox
    $gb.Text = Convert-UiText $Texts.GbSettings
    $gb.Location = New-Object System.Drawing.Point(20, 174)
    $gb.Size = New-Object System.Drawing.Size(924, 190)
    $form.Controls.Add($gb)

    $lblDom = New-Object System.Windows.Forms.Label
    $lblDom.Text = Convert-UiText $Texts.LblDomain
    $lblDom.Location = New-Object System.Drawing.Point(18, 34)
    $lblDom.AutoSize = $true
    $gb.Controls.Add($lblDom)

    $txtDomain = New-Object System.Windows.Forms.TextBox
    $txtDomain.Location = New-Object System.Drawing.Point(18, 58)
    $txtDomain.Size = New-Object System.Drawing.Size(300, 27)
    $gb.Controls.Add($txtDomain)
    $toolTip.SetToolTip($txtDomain, (Convert-UiText $Texts.TipDomain))

    $chkNever = New-Object System.Windows.Forms.CheckBox
    $chkNever.Text = Convert-UiText $Texts.ChkNever
    $chkNever.Location = New-Object System.Drawing.Point(350, 60)
    $chkNever.AutoSize = $true
    $gb.Controls.Add($chkNever)

    $lblOU = New-Object System.Windows.Forms.Label
    $lblOU.Text = Convert-UiText $Texts.LblOu
    $lblOU.Location = New-Object System.Drawing.Point(18, 100)
    $lblOU.AutoSize = $true
    $gb.Controls.Add($lblOU)

    $txtOU = New-Object System.Windows.Forms.TextBox
    $txtOU.Location = New-Object System.Drawing.Point(18, 124)
    $txtOU.Size = New-Object System.Drawing.Size(720, 27)
    $gb.Controls.Add($txtOU)

    $btnOU = New-Object System.Windows.Forms.Button
    $btnOU.Text = Convert-UiText $Texts.BtnOu
    $btnOU.Location = New-Object System.Drawing.Point(756, 121)
    $btnOU.Size = New-Object System.Drawing.Size(148, 34)
    $btnOU.BackColor = [System.Drawing.Color]::FromArgb(243, 247, 255)
    $btnOU.FlatStyle = 'Flat'
    $gb.Controls.Add($btnOU)

    $lblGroups = New-Object System.Windows.Forms.Label
    $lblGroups.Text = Convert-UiText $Texts.LblGroups
    $lblGroups.Location = New-Object System.Drawing.Point(18, 160)
    $lblGroups.AutoSize = $true
    $gb.Controls.Add($lblGroups)

    $txtGroups = New-Object System.Windows.Forms.TextBox
    $txtGroups.Location = New-Object System.Drawing.Point(320, 157)
    $txtGroups.Size = New-Object System.Drawing.Size(418, 27)
    $gb.Controls.Add($txtGroups)
    $toolTip.SetToolTip($txtGroups, (Convert-UiText $Texts.TipGroups))

    $btnGroups = New-Object System.Windows.Forms.Button
    $btnGroups.Text = Convert-UiText $Texts.BtnGroups
    $btnGroups.Location = New-Object System.Drawing.Point(756, 154)
    $btnGroups.Size = New-Object System.Drawing.Size(148, 34)
    $btnGroups.BackColor = [System.Drawing.Color]::FromArgb(243, 247, 255)
    $btnGroups.FlatStyle = 'Flat'
    $gb.Controls.Add($btnGroups)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = Convert-UiText $Texts.BtnRun
    $btnRun.Location = New-Object System.Drawing.Point(20, 386)
    $btnRun.Size = New-Object System.Drawing.Size(924, 50)
    $btnRun.BackColor = [System.Drawing.Color]::FromArgb(208, 245, 217)
    $btnRun.FlatStyle = 'Flat'
    $btnRun.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($btnRun)

    $lblLog = New-Object System.Windows.Forms.Label
    $lblLog.Text = Convert-UiText $Texts.LblLog
    $lblLog.Location = New-Object System.Drawing.Point(20, 452)
    $lblLog.AutoSize = $true
    $lblLog.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblLog)

    $txtLog = New-Object System.Windows.Forms.TextBox
    $txtLog.Location = New-Object System.Drawing.Point(20, 478)
    $txtLog.Size = New-Object System.Drawing.Size(924, 182)
    $txtLog.Multiline = $true
    $txtLog.ReadOnly = $true
    $txtLog.ScrollBars = "Vertical"
    $txtLog.Font = New-Object System.Drawing.Font("Consolas", 9.5)
    $txtLog.BackColor = [System.Drawing.Color]::FromArgb(249, 251, 255)
    $form.Controls.Add($txtLog)

    return @{
        Form      = $form
        ToolTip   = $toolTip
        TxtExcel  = $txtExcel
        BtnExcel  = $btnExcel
        TxtPdf    = $txtPdf
        BtnPdf    = $btnPdf
        TxtDomain = $txtDomain
        ChkNever  = $chkNever
        TxtOU     = $txtOU
        BtnOU     = $btnOU
        TxtGroups = $txtGroups
        BtnGroups = $btnGroups
        BtnRun    = $btnRun
        TxtLog    = $txtLog
    }
}
