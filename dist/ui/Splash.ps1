function Show-Splash {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $splash = New-Object System.Windows.Forms.Form
    $splash.Text = "AD User Creator"
    $splash.Size = New-Object System.Drawing.Size(420, 140)
    $splash.StartPosition = "CenterScreen"
    $splash.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $splash.MaximizeBox = $false
    $splash.MinimizeBox = $false
    $splash.TopMost = $true

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Dock = [System.Windows.Forms.DockStyle]::Fill
    $lbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $lbl.Text = "Завантаження модулів...`r`nПочекайте"
    $splash.Controls.Add($lbl)

    # правильна підписка на подію
    $splash.add_Shown({ $splash.Refresh() })

    $splash.Show()
    [System.Windows.Forms.Application]::DoEvents()

    return $splash
}
