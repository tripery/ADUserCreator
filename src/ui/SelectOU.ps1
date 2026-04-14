function Select-OU {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select target OU"
    $form.Size = New-Object System.Drawing.Size(900, 720)
    $form.MinimumSize = New-Object System.Drawing.Size(760, 620)
    $form.StartPosition = "CenterScreen"
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 252)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $header = New-Object System.Windows.Forms.Panel
    $header.Dock = "Top"
    $header.Height = 78
    $header.BackColor = [System.Drawing.Color]::FromArgb(24, 30, 43)
    $form.Controls.Add($header)

    $title = New-Object System.Windows.Forms.Label
    $title.Text = "Target OU"
    $title.Location = New-Object System.Drawing.Point(20, 14)
    $title.AutoSize = $true
    $title.ForeColor = [System.Drawing.Color]::White
    $title.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $header.Controls.Add($title)

    $subtitle = New-Object System.Windows.Forms.Label
    $subtitle.Text = "Choose the organizational unit for new users."
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
    $txtSearch.Size = New-Object System.Drawing.Size(840, 25)
    $txtSearch.Anchor = "Top,Left,Right"
    $body.Controls.Add($txtSearch)

    $tree = New-Object System.Windows.Forms.TreeView
    $tree.Location = New-Object System.Drawing.Point(6, 60)
    $tree.Size = New-Object System.Drawing.Size(840, 500)
    $tree.Anchor = "Top,Bottom,Left,Right"
    $tree.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $tree.HideSelection = $false
    $body.Controls.Add($tree)

    $status = New-Object System.Windows.Forms.Label
    $status.Location = New-Object System.Drawing.Point(6, 570)
    $status.Size = New-Object System.Drawing.Size(840, 24)
    $status.Anchor = "Bottom,Left,Right"
    $status.TextAlign = "MiddleLeft"
    $status.ForeColor = [System.Drawing.Color]::FromArgb(70, 78, 90)
    $body.Controls.Add($status)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Size = New-Object System.Drawing.Size(120, 36)
    $btnCancel.Location = New-Object System.Drawing.Point(586, 602)
    $btnCancel.Anchor = "Bottom,Right"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $btnCancel.FlatStyle = "Flat"
    $btnCancel.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(170, 180, 195)
    $body.Controls.Add($btnCancel)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Select OU"
    $btnOK.Enabled = $false
    $btnOK.Size = New-Object System.Drawing.Size(140, 36)
    $btnOK.Location = New-Object System.Drawing.Point(712, 602)
    $btnOK.Anchor = "Bottom,Right"
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnOK.BackColor = [System.Drawing.Color]::FromArgb(228, 238, 255)
    $btnOK.FlatStyle = "Flat"
    $btnOK.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(140, 170, 220)
    $body.Controls.Add($btnOK)

    $nodeMap = @{}

    function Find-NodeByText {
        param(
            [System.Windows.Forms.TreeNodeCollection]$Nodes,
            [string]$Query
        )

        foreach ($node in $Nodes) {
            if ($node.Text -like "*$Query*" -or $node.Tag -like "*$Query*") {
                return $node
            }

            $found = Find-NodeByText -Nodes $node.Nodes -Query $Query
            if ($found) { return $found }
        }

        return $null
    }

    try {
        $domainDN = (Get-ADDomain).DistinguishedName

        $rootNode = $tree.Nodes.Add("OU: $domainDN")
        $rootNode.Tag = $domainDN
        $rootNode.NodeFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

        $allObjects = Get-ADObject -LDAPFilter "(|(objectClass=organizationalUnit)(objectClass=container))" -SearchBase $domainDN -ErrorAction Stop

        $nodeMap[$domainDN] = $rootNode

        foreach ($obj in ($allObjects | Sort-Object { $_.DistinguishedName.Length })) {
            $dn = $obj.DistinguishedName
            if ($dn -eq $domainDN) { continue }

            $parentDN = ($dn -split ',', 2)[1]
            $parentNode = $nodeMap[$parentDN]
            if (-not $parentNode) { $parentNode = $rootNode }

            $newNode = $parentNode.Nodes.Add($obj.Name)
            $newNode.Tag = $dn
            $nodeMap[$dn] = $newNode
        }

        $rootNode.Expand()
        $tree.SelectedNode = $rootNode
        $status.Text = "Ready. Select an OU or search by name."
        $btnOK.Enabled = $true
    }
    catch {
        $status.Text = "ERROR: $($_.Exception.Message)"
        $status.ForeColor = [System.Drawing.Color]::Red
    }

    $tree.Add_AfterSelect({
        if ($tree.SelectedNode -and $tree.SelectedNode.Tag) {
            $btnOK.Enabled = $true
            $status.Text = "Selected: $($tree.SelectedNode.Text)"
        }
    })

    $txtSearch.Add_TextChanged({
        $query = $txtSearch.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($query)) {
            if ($tree.Nodes.Count -gt 0) {
                $tree.SelectedNode = $tree.Nodes[0]
            }
            return
        }

        $foundNode = Find-NodeByText -Nodes $tree.Nodes -Query $query
        if ($foundNode) {
            $tree.SelectedNode = $foundNode
            $foundNode.EnsureVisible()
            $status.Text = "Selected: $($foundNode.Text)"
        }
    })

    $form.AcceptButton = $btnOK
    $form.CancelButton = $btnCancel

    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and $tree.SelectedNode -and $tree.SelectedNode.Tag) {
        return $tree.SelectedNode.Tag
    }

    return $null
}
