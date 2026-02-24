function Select-OU {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Виберіть OU"
    $form.Size = New-Object System.Drawing.Size(800,680)
    $form.StartPosition = "CenterScreen"

    $tree = New-Object System.Windows.Forms.TreeView
    $tree.Dock = "Fill"
    $tree.Font = New-Object System.Drawing.Font("Segoe UI",10)
    $form.Controls.Add($tree)

    $status = New-Object System.Windows.Forms.Label
    $status.Dock = "Bottom"
    $status.Height = 30
    $status.TextAlign = "MiddleCenter"
    $form.Controls.Add($status)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Enabled = $false
    $btnOK.Dock = "Bottom"
    $btnOK.Height = 40
    $btnOK.DialogResult = "OK"
    $form.Controls.Add($btnOK)

    try {
        $domainDN = (Get-ADDomain).DistinguishedName

        $rootNode = $tree.Nodes.Add("OU: $domainDN")
        $rootNode.Tag = $domainDN
        $rootNode.NodeFont = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)

       $allObjects = Get-ADObject -LDAPFilter "(|(objectClass=organizationalUnit)(objectClass=container))" -SearchBase $domainDN -ErrorAction Stop

        $nodeMap = @{}
        $nodeMap[$domainDN] = $rootNode

        foreach ($obj in ($allObjects | Sort-Object {$_.DistinguishedName.Length})) {
            $dn = $obj.DistinguishedName
            if ($dn -eq $domainDN) { continue }

            $parentDN = ($dn -split ',',2)[1]
            $parentNode = $nodeMap[$parentDN]
            if (-not $parentNode) { $parentNode = $rootNode }

            $newNode = $parentNode.Nodes.Add($obj.Name)
            $newNode.Tag = $dn
            $nodeMap[$dn] = $newNode
        }

        $rootNode.Expand()
        $status.Text = "Готово! Виберіть OU і натисніть OK"
        $btnOK.Enabled = $true
    }
    catch {
        $status.Text = "ПОМИЛКА: $($_.Exception.Message)"
        $status.ForeColor = "Red"
    }

    $tree.Add_AfterSelect({
        if ($tree.SelectedNode.Tag) {
            $btnOK.Enabled = $true
            $status.Text = "Вибрано: $($tree.SelectedNode.Text)"
        }
    })

    if ($form.ShowDialog() -eq "OK" -and $tree.SelectedNode.Tag) {
        return $tree.SelectedNode.Tag
    }
    return $null
}
