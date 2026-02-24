function Select-Groups {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Вибір груп AD"
    $form.Size = New-Object System.Drawing.Size(520, 420)
    $form.StartPosition = "CenterScreen"

    $listbox = New-Object System.Windows.Forms.ListBox
    $listbox.Dock = "Fill"
    $listbox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $listbox.SelectionMode = "MultiExtended"
    $form.Controls.Add($listbox)

    $status = New-Object System.Windows.Forms.Label
    $status.Dock = "Bottom"
    $status.Height = 30
    $status.TextAlign = "MiddleCenter"
    $form.Controls.Add($status)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Dock = "Bottom"
    $btnOK.Height = 40
    $btnOK.DialogResult = "OK"
    $form.Controls.Add($btnOK)

    $status.Text = "Завантаження груп..."

    try {
        $groups = Get-ADGroup -Filter { GroupCategory -eq 'Security' } -Properties Name, SamAccountName, GroupScope |
            Select-Object Name, SamAccountName, GroupScope

        foreach ($g in ($groups | Sort-Object Name)) {
            # показуємо Name (Scope) — але повернемо SamAccountName
            $listbox.Items.Add([pscustomobject]@{ Text="$($g.Name) ($($g.GroupScope))"; Sam=$g.SamAccountName }) | Out-Null
        }

        # рендер тексту
        $listbox.DisplayMember = "Text"

        $status.Text = "Виберіть одну або кілька груп і натисніть OK."
    }
    catch {
        $status.Text = "ПОМИЛКА: $($_.Exception.Message)"
        $status.ForeColor = "Red"
    }

    if ($form.ShowDialog() -eq "OK") {
        $selected = @()
        foreach ($it in $listbox.SelectedItems) {
            $selected += $it.Sam
        }
        return $selected
    }

    return $null
}
