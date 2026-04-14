function Select-Groups {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select AD groups"
    $form.Size = New-Object System.Drawing.Size(760, 620)
    $form.MinimumSize = New-Object System.Drawing.Size(660, 520)
    $form.StartPosition = "CenterScreen"
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 252)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $header = New-Object System.Windows.Forms.Panel
    $header.Dock = "Top"
    $header.Height = 78
    $header.BackColor = [System.Drawing.Color]::FromArgb(24, 30, 43)
    $form.Controls.Add($header)

    $title = New-Object System.Windows.Forms.Label
    $title.Text = "AD groups"
    $title.Location = New-Object System.Drawing.Point(20, 14)
    $title.AutoSize = $true
    $title.ForeColor = [System.Drawing.Color]::White
    $title.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $header.Controls.Add($title)

    $subtitle = New-Object System.Windows.Forms.Label
    $subtitle.Text = "Select one or more security groups for new users."
    $subtitle.Location = New-Object System.Drawing.Point(22, 44)
    $subtitle.AutoSize = $true
    $subtitle.ForeColor = [System.Drawing.Color]::FromArgb(205, 214, 224)
    $subtitle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $header.Controls.Add($subtitle)

    $body = New-Object System.Windows.Forms.Panel
    $body.Dock = "Fill"
    $body.Padding = New-Object System.Windows.Forms.Padding(16, 14, 16, 14)
    $form.Controls.Add($body)

    $searchLabel = New-Object System.Windows.Forms.Label
    $searchLabel.Text = "Search"
    $searchLabel.Location = New-Object System.Drawing.Point(6, 4)
    $searchLabel.AutoSize = $true
    $searchLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $body.Controls.Add($searchLabel)

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(6, 24)
    $txtSearch.Size = New-Object System.Drawing.Size(700, 25)
    $txtSearch.Anchor = "Top,Left,Right"
    $body.Controls.Add($txtSearch)

    $listbox = New-Object System.Windows.Forms.ListBox
    $listbox.Location = New-Object System.Drawing.Point(6, 60)
    $listbox.Size = New-Object System.Drawing.Size(700, 410)
    $listbox.Anchor = "Top,Bottom,Left,Right"
    $listbox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $listbox.SelectionMode = "MultiExtended"
    $body.Controls.Add($listbox)

    $status = New-Object System.Windows.Forms.Label
    $status.Location = New-Object System.Drawing.Point(6, 482)
    $status.Size = New-Object System.Drawing.Size(700, 24)
    $status.Anchor = "Bottom,Left,Right"
    $status.TextAlign = "MiddleLeft"
    $status.ForeColor = [System.Drawing.Color]::FromArgb(70, 78, 90)
    $body.Controls.Add($status)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Size = New-Object System.Drawing.Size(120, 36)
    $btnCancel.Location = New-Object System.Drawing.Point(460, 516)
    $btnCancel.Anchor = "Bottom,Right"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $btnCancel.FlatStyle = "Flat"
    $btnCancel.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(170, 180, 195)
    $body.Controls.Add($btnCancel)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Select groups"
    $btnOK.Size = New-Object System.Drawing.Size(140, 36)
    $btnOK.Location = New-Object System.Drawing.Point(566, 516)
    $btnOK.Anchor = "Bottom,Right"
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnOK.BackColor = [System.Drawing.Color]::FromArgb(228, 238, 255)
    $btnOK.FlatStyle = "Flat"
    $btnOK.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(140, 170, 220)
    $body.Controls.Add($btnOK)

    $allItems = New-Object System.Collections.ArrayList
    $status.Text = "Loading groups..."

    try {
        $groups = Get-ADGroup -Filter { GroupCategory -eq 'Security' } -Properties Name, SamAccountName, GroupScope |
            Select-Object Name, SamAccountName, GroupScope

        foreach ($g in ($groups | Sort-Object Name)) {
            $item = [pscustomobject]@{
                Text = "$($g.Name) ($($g.GroupScope))"
                Sam  = $g.SamAccountName
            }
            [void]$allItems.Add($item)
        }

        $listbox.DisplayMember = "Text"
        $listbox.Items.Clear()
        foreach ($item in $allItems) {
            [void]$listbox.Items.Add($item)
        }

        $status.Text = "Ready. Select one or more groups."
    }
    catch {
        $status.Text = "ERROR: $($_.Exception.Message)"
        $status.ForeColor = [System.Drawing.Color]::Red
    }

    $txtSearch.Add_TextChanged({
        $selectedSam = @()
        foreach ($it in $listbox.SelectedItems) {
            $selectedSam += $it.Sam
        }

        $listbox.Items.Clear()
        $query = $txtSearch.Text.Trim().ToLowerInvariant()

        foreach ($item in $allItems) {
            if ([string]::IsNullOrWhiteSpace($query) -or $item.Text.ToLowerInvariant().Contains($query) -or $item.Sam.ToLowerInvariant().Contains($query)) {
                [void]$listbox.Items.Add($item)
            }
        }

        for ($i = 0; $i -lt $listbox.Items.Count; $i++) {
            if ($selectedSam -contains $listbox.Items[$i].Sam) {
                $listbox.SetSelected($i, $true)
            }
        }

        $status.Text = "Shown groups: $($listbox.Items.Count)"
    })

    $listbox.Add_SelectedIndexChanged({
        $status.Text = "Selected groups: $($listbox.SelectedItems.Count)"
    })

    $form.AcceptButton = $btnOK
    $form.CancelButton = $btnCancel

    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $selected = @()
        foreach ($it in $listbox.SelectedItems) {
            $selected += $it.Sam
        }
        return $selected
    }

    return $null
}
