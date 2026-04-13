$formDir = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $formDir 'Form.Localization.ps1')
. (Join-Path $formDir 'Form.Layout.ps1')
. (Join-Path $formDir 'Form.Events.ps1')

<<<<<<< Updated upstream
    function Get-UiText([string]$Encoded) {
        return [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($Encoded))
    }

    $ui = @{
        FormTitle            = Get-UiText '0KHRgtCy0L7RgNC10L3QvdGPINC60L7RgNC40YHRgtGD0LLQsNGH0ZbQsiBBRCAoRXhjZWwgLT4gQUQp'
        ExcelLabel           = Get-UiText 'MS4gRXhjZWwt0YTQsNC50Ls='
        ExcelHint            = Get-UiText '0J7Rh9GW0LrRg9GO0YLRjNGB0Y8g0LrQvtC70L7QvdC60LggKi54bHN4OiBWc3R1cG55aywgU3RydWt0dXJueWkgcGlkcm96ZGlsLg=='
        ChooseFile           = Get-UiText '0JLQuNCx0YDQsNGC0Lgg0YTQsNC50Ls='
        PdfFolderLabel       = Get-UiText '0J/QsNC/0LrQsCDQtNC70Y8g0LfQsdC10YDQtdC20LXQvdC90Y8gUERGOg=='
        PdfFolderTip         = Get-UiText '0J7QsdC10YDRltGC0Ywg0L/QsNC/0LrRgywg0LrRg9C00Lgg0LHRg9C00YPRgtGMINC30LHQtdGA0LXQttC10L3RliBQREYg0Lcg0L7QsdC70ZbQutC+0LLQuNC80Lgg0LTQsNC90LjQvNC4INGC0LAgSlNPTi3QvNC10YLQsNC00LDQvdGWLg=='
        ChooseFolder         = Get-UiText '0JLQuNCx0YDQsNGC0Lgg0L/QsNC/0LrRgy4uLg=='
        AdSettings           = Get-UiText 'Mi4g0J3QsNC70LDRiNGC0YPQstCw0L3QvdGPIEFE'
        DomainLabel          = Get-UiText '0JTQvtC80LXQvSDQtNC70Y8gVVBOL9C/0L7RiNGC0Lg6'
        DomainTip            = Get-UiText '0J/RgNC40LrQu9Cw0LQ6IGRvbm51LmVkdS51YQ=='
        PasswordNever        = Get-UiText '0J/QsNGA0L7Qu9GMINC90LUg0LfQsNC60ZbQvdGH0YPRlNGC0YzRgdGP'
        TargetOu             = Get-UiText '0KbRltC70YzQvtCy0LAgT1U6'
        ChooseOu             = Get-UiText '0JLQuNCx0YDQsNGC0LggT1UuLi4='
        GroupsLabel          = Get-UiText '0JPRgNGD0L/QuCAoU2FtQWNjb3VudE5hbWUg0YfQtdGA0LXQtyDQutC+0LzRgyk6'
        GroupsTip            = Get-UiText '0JLQstC10LTRltGC0Ywg0LPRgNGD0L/QuCDRh9C10YDQtdC3INC60L7QvNGDINCw0LHQviDRgdC60L7RgNC40YHRgtCw0LnRgtC10YHRjyDQutC90L7Qv9C60L7RjiDQktC40LHRgNCw0YLQuCDQs9GA0YPQv9C4Li4u'
        ChooseGroups         = Get-UiText '0JLQuNCx0YDQsNGC0Lgg0LPRgNGD0L/QuC4uLg=='
        CreateUsers          = Get-UiText '0KHQotCS0J7QoNCY0KLQmCDQmtCe0KDQmNCh0KLQo9CS0JDQp9CG0JI='
        LogLabel             = Get-UiText '0JbRg9GA0L3QsNC7INCy0LjQutC+0L3QsNC90L3Rjzo='
        OpeningExcel         = Get-UiText '0JLRltC00LrRgNC40YLRgtGPIEV4Y2VsOg=='
        UsingSheet           = Get-UiText '0JLQuNC60L7RgNC40YHRgtC+0LLRg9GU0YLRjNGB0Y8g0LvQuNGB0YI6'
        Rows                 = Get-UiText '0KDRj9C00LrRltCyOg=='
        ExcelErrorLog        = Get-UiText '0J/QvtC80LjQu9C60LAgRXhjZWw6'
        ExcelErrorTitle      = Get-UiText '0J/QvtC80LjQu9C60LAgRXhjZWw='
        OuSelected           = Get-UiText 'T1Ug0LLQuNCx0YDQsNC90L46'
        OuPickerError        = Get-UiText '0J/QvtC80LjQu9C60LAg0LLQuNCx0L7RgNGDIE9VOg=='
        GroupsSelected       = Get-UiText '0JPRgNGD0L/QuCDQstC40LHRgNCw0L3Qvjo='
        GroupSelectionError  = Get-UiText '0J/QvtC80LjQu9C60LAg0LLQuNCx0L7RgNGDINCz0YDRg9C/Og=='
        PdfDialogDescription = Get-UiText '0J7QsdC10YDRltGC0Ywg0L/QsNC/0LrRgyDQtNC70Y8gUERGLdGE0LDQudC70ZbQsiDQtyDQv9Cw0YDQvtC70Y/QvNC4'
        PdfFolderSelected    = Get-UiText '0J/QsNC/0LrRgyDQtNC70Y8gUERGINCy0LjQsdGA0LDQvdC+Og=='
        ChooseExcelFirst     = Get-UiText '0KHQv9C+0YfQsNGC0LrRgyDQstC40LHQtdGA0ZbRgtGMIEV4Y2VsLdGE0LDQudC7Lg=='
        NoDataTitle          = Get-UiText '0J3QtdC80LDRlCDQtNCw0L3QuNGF'
        ChooseOuFirst        = Get-UiText '0JLQuNCx0LXRgNGW0YLRjCBPVS4='
        NoOuTitle            = Get-UiText '0J3QtdC80LDRlCBPVQ=='
        EnterDomain          = Get-UiText '0JLQstC10LTRltGC0Ywg0LTQvtC80LXQvSDQtNC70Y8gVVBOL9C/0L7RiNGC0Lgu'
        NoDomainTitle        = Get-UiText '0J3QtdC80LDRlCDQtNC+0LzQtdC90YM='
        StartLogPrefix       = Get-UiText '0KHQotCQ0KDQojog0YHRgtCy0L7RgNC10L3QvdGPINC60L7RgNC40YHRgtGD0LLQsNGH0ZbQsi4gT1U9'
        DomainPrefix         = Get-UiText 'LCDQlNC+0LzQtdC9PQ=='
        GroupsPrefix         = Get-UiText 'LCDQk9GA0YPQv9C4PQ=='
        PdfSaved             = Get-UiText 'UERGINC3INC/0LDRgNC+0LvRj9C80Lgg0LfQsdC10YDQtdC20LXQvdC+Og=='
        PdfError             = Get-UiText '0J/QvtC80LjQu9C60LAg0LPQtdC90LXRgNCw0YbRltGXIFBERiDQtyDQv9Cw0YDQvtC70Y/QvNC4Og=='
    }

    $colorWindow = [System.Drawing.Color]::FromArgb(245, 248, 252)
    $colorPanel = [System.Drawing.Color]::White
    $colorPanelAlt = [System.Drawing.Color]::FromArgb(249, 251, 255)
    $colorBorder = [System.Drawing.Color]::FromArgb(214, 223, 235)
    $colorText = [System.Drawing.Color]::FromArgb(30, 41, 59)
    $colorMuted = [System.Drawing.Color]::FromArgb(100, 116, 139)
    $colorAccent = [System.Drawing.Color]::FromArgb(29, 78, 216)
    $colorAccentSoft = [System.Drawing.Color]::FromArgb(229, 239, 255)
    $colorSuccess = [System.Drawing.Color]::FromArgb(22, 101, 52)
    $colorSuccessSoft = [System.Drawing.Color]::FromArgb(220, 252, 231)

    $fontBase = New-Object System.Drawing.Font("Segoe UI", 10)
    $fontLabel = New-Object System.Drawing.Font("Segoe UI", 10)
    $fontSection = New-Object System.Drawing.Font("Segoe UI Semibold", 10.5, [System.Drawing.FontStyle]::Bold)
    $fontRun = New-Object System.Drawing.Font("Segoe UI Semibold", 12, [System.Drawing.FontStyle]::Bold)
    $fontLog = New-Object System.Drawing.Font("Consolas", 9.5)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $ui.FormTitle
    $form.Size = New-Object System.Drawing.Size(980, 610)
    $form.MinimumSize = New-Object System.Drawing.Size(980, 610)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $form.MaximizeBox = $false
    $form.BackColor = $colorWindow
    $form.ForeColor = $colorText
    $form.Font = $fontBase

    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.AutoPopDelay = 20000
    $toolTip.InitialDelay = 500
    $toolTip.ReshowDelay = 200
    $toolTip.ShowAlways = $true

    $lblExcel = New-Object System.Windows.Forms.Label
    $lblExcel.Text = $ui.ExcelLabel
    $lblExcel.Location = New-Object System.Drawing.Point(20, 20)
    $lblExcel.AutoSize = $true
    $lblExcel.Font = $fontSection
    $form.Controls.Add($lblExcel)

    $lblExcelHint = New-Object System.Windows.Forms.Label
    $lblExcelHint.Text = $ui.ExcelHint
    $lblExcelHint.Location = New-Object System.Drawing.Point(20, 44)
    $lblExcelHint.AutoSize = $true
    $lblExcelHint.ForeColor = $colorMuted
    $form.Controls.Add($lblExcelHint)

    $txtExcel = New-Object System.Windows.Forms.TextBox
    $txtExcel.Location = New-Object System.Drawing.Point(20, 70)
    $txtExcel.Size = New-Object System.Drawing.Size(700, 31)
    $txtExcel.BackColor = $colorPanel
    $txtExcel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $form.Controls.Add($txtExcel)

    $btnExcel = New-Object System.Windows.Forms.Button
    $btnExcel.Text = $ui.ChooseFile
    $btnExcel.Location = New-Object System.Drawing.Point(738, 67)
    $btnExcel.Size = New-Object System.Drawing.Size(202, 36)
    $btnExcel.BackColor = $colorAccentSoft
    $btnExcel.ForeColor = $colorAccent
    $btnExcel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnExcel.FlatAppearance.BorderColor = $colorBorder
    $btnExcel.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(217, 232, 255)
    $form.Controls.Add($btnExcel)

    $lblPdf = New-Object System.Windows.Forms.Label
    $lblPdf.Text = $ui.PdfFolderLabel
    $lblPdf.Location = New-Object System.Drawing.Point(20, 116)
    $lblPdf.AutoSize = $true
    $lblPdf.Font = $fontSection
    $form.Controls.Add($lblPdf)

    $txtPdfFolder = New-Object System.Windows.Forms.TextBox
    $txtPdfFolder.Location = New-Object System.Drawing.Point(20, 142)
    $txtPdfFolder.Size = New-Object System.Drawing.Size(700, 31)
    $txtPdfFolder.BackColor = $colorPanel
    $txtPdfFolder.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $txtPdfFolder.Text = $script:PasswordLogsRoot
    $form.Controls.Add($txtPdfFolder)
    $toolTip.SetToolTip($txtPdfFolder, $ui.PdfFolderTip)

    $btnPdfFolder = New-Object System.Windows.Forms.Button
    $btnPdfFolder.Text = $ui.ChooseFolder
    $btnPdfFolder.Location = New-Object System.Drawing.Point(738, 139)
    $btnPdfFolder.Size = New-Object System.Drawing.Size(202, 36)
    $btnPdfFolder.BackColor = $colorPanel
    $btnPdfFolder.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnPdfFolder.FlatAppearance.BorderColor = $colorBorder
    $form.Controls.Add($btnPdfFolder)

    $gb = New-Object System.Windows.Forms.GroupBox
    $gb.Text = $ui.AdSettings
    $gb.Location = New-Object System.Drawing.Point(20, 200)
    $gb.Size = New-Object System.Drawing.Size(920, 182)
    $gb.BackColor = $colorPanelAlt
    $gb.ForeColor = $colorText
    $gb.Font = $fontSection
    $form.Controls.Add($gb)

    $lblDom = New-Object System.Windows.Forms.Label
    $lblDom.Text = $ui.DomainLabel
    $lblDom.Location = New-Object System.Drawing.Point(15, 32)
    $lblDom.AutoSize = $true
    $lblDom.Font = $fontLabel
    $gb.Controls.Add($lblDom)

    $txtDomain = New-Object System.Windows.Forms.TextBox
    $txtDomain.Location = New-Object System.Drawing.Point(210, 28)
    $txtDomain.Size = New-Object System.Drawing.Size(260, 31)
    $txtDomain.BackColor = $colorPanel
    $txtDomain.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $gb.Controls.Add($txtDomain)
    $toolTip.SetToolTip($txtDomain, $ui.DomainTip)

    try { $txtDomain.Text = (Get-ADDomain).DNSRoot } catch {}

    $chkNever = New-Object System.Windows.Forms.CheckBox
    $chkNever.Text = $ui.PasswordNever
    $chkNever.Location = New-Object System.Drawing.Point(500, 31)
    $chkNever.AutoSize = $true
    $chkNever.Font = $fontLabel
    $gb.Controls.Add($chkNever)

    $lblOU = New-Object System.Windows.Forms.Label
    $lblOU.Text = $ui.TargetOu
    $lblOU.Location = New-Object System.Drawing.Point(15, 74)
    $lblOU.AutoSize = $true
    $lblOU.Font = $fontLabel
    $gb.Controls.Add($lblOU)

    $txtOU = New-Object System.Windows.Forms.TextBox
    $txtOU.Location = New-Object System.Drawing.Point(210, 70)
    $txtOU.Size = New-Object System.Drawing.Size(520, 31)
    $txtOU.BackColor = $colorPanel
    $txtOU.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $gb.Controls.Add($txtOU)

    $btnOU = New-Object System.Windows.Forms.Button
    $btnOU.Text = $ui.ChooseOu
    $btnOU.Location = New-Object System.Drawing.Point(745, 67)
    $btnOU.Size = New-Object System.Drawing.Size(160, 36)
    $btnOU.BackColor = $colorPanel
    $btnOU.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOU.FlatAppearance.BorderColor = $colorBorder
    $gb.Controls.Add($btnOU)

    $lblGroups = New-Object System.Windows.Forms.Label
    $lblGroups.Text = $ui.GroupsLabel
    $lblGroups.Location = New-Object System.Drawing.Point(15, 116)
    $lblGroups.AutoSize = $true
    $lblGroups.Font = $fontLabel
    $gb.Controls.Add($lblGroups)

    $txtGroups = New-Object System.Windows.Forms.TextBox
    $txtGroups.Location = New-Object System.Drawing.Point(15, 142)
    $txtGroups.Size = New-Object System.Drawing.Size(715, 31)
    $txtGroups.BackColor = $colorPanel
    $txtGroups.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $gb.Controls.Add($txtGroups)
    $toolTip.SetToolTip($txtGroups, $ui.GroupsTip)

    $btnGroups = New-Object System.Windows.Forms.Button
    $btnGroups.Text = $ui.ChooseGroups
    $btnGroups.Location = New-Object System.Drawing.Point(745, 139)
    $btnGroups.Size = New-Object System.Drawing.Size(160, 36)
    $btnGroups.BackColor = $colorPanel
    $btnGroups.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnGroups.FlatAppearance.BorderColor = $colorBorder
    $gb.Controls.Add($btnGroups)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = $ui.CreateUsers
    $btnRun.Location = New-Object System.Drawing.Point(20, 400)
    $btnRun.Size = New-Object System.Drawing.Size(920, 56)
    $btnRun.BackColor = $colorSuccessSoft
    $btnRun.ForeColor = $colorSuccess
    $btnRun.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnRun.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(167, 243, 208)
    $btnRun.FlatAppearance.BorderSize = 1
    $btnRun.Font = $fontRun
    $form.Controls.Add($btnRun)

    $lblLog = New-Object System.Windows.Forms.Label
    $lblLog.Text = $ui.LogLabel
    $lblLog.Location = New-Object System.Drawing.Point(20, 473)
    $lblLog.AutoSize = $true
    $lblLog.Font = $fontSection
    $form.Controls.Add($lblLog)

    $txtLog = New-Object System.Windows.Forms.TextBox
    $txtLog.Location = New-Object System.Drawing.Point(20, 499)
    $txtLog.Size = New-Object System.Drawing.Size(920, 110)
    $txtLog.Multiline = $true
    $txtLog.ReadOnly = $true
    $txtLog.ScrollBars = "Vertical"
    $txtLog.BackColor = $colorPanel
    $txtLog.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $txtLog.Font = $fontLog
    $form.Controls.Add($txtLog)

    Set-LogTarget -TextBox $txtLog

    $script:LoadedUsers = $null
    $script:LoadedSheet = $null

    $btnExcel.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
        if ($dlg.ShowDialog() -ne "OK") { return }

        $txtExcel.Text = $dlg.FileName
        $script:LoadedUsers = $null
        $script:LoadedSheet = $null

        try {
            Write-Log "$($ui.OpeningExcel) $($txtExcel.Text)" "INFO"
            $res = Import-UsersFromExcelSmart -Path $txtExcel.Text
            $script:LoadedUsers = $res.Users
            $script:LoadedSheet = $res.Sheet

            Write-Log "$($ui.UsingSheet) $($script:LoadedSheet). $($ui.Rows) $($script:LoadedUsers.Count)" "OK"

        }
        catch {
            Write-Log "$($ui.ExcelErrorLog) $($_.Exception.Message)" "ERROR"
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, $ui.ExcelErrorTitle, "OK", "Error") | Out-Null
        }
    })

    $btnOU.Add_Click({
        try {
            $ou = Select-OU
            if ($ou) { $txtOU.Text = $ou; Write-Log "$($ui.OuSelected) $ou" "OK" }
        } catch {
            Write-Log "$($ui.OuPickerError) $($_.Exception.Message)" "ERROR"
        }
    })

    $btnGroups.Add_Click({
        try {
            $sel = Select-Groups
            if ($sel -and $sel.Count -gt 0) {
                $txtGroups.Text = ($sel -join ",")
                Write-Log "$($ui.GroupsSelected) $($sel -join ', ')" "OK"
            }
        } catch {
            Write-Log "$($ui.GroupSelectionError) $($_.Exception.Message)" "ERROR"
        }
    })

    $btnPdfFolder.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = $ui.PdfDialogDescription
        if (-not [string]::IsNullOrWhiteSpace($txtPdfFolder.Text) -and (Test-Path -LiteralPath $txtPdfFolder.Text)) {
            $dlg.SelectedPath = $txtPdfFolder.Text
        }
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtPdfFolder.Text = $dlg.SelectedPath
            Write-Log "$($ui.PdfFolderSelected) $($dlg.SelectedPath)" "OK"
        }
    })

    $btnRun.Add_Click({
        if (-not $script:LoadedUsers) { [System.Windows.Forms.MessageBox]::Show($ui.ChooseExcelFirst, $ui.NoDataTitle, "OK", "Warning") | Out-Null; return }
        if ([string]::IsNullOrWhiteSpace($txtOU.Text)) { [System.Windows.Forms.MessageBox]::Show($ui.ChooseOuFirst, $ui.NoOuTitle, "OK", "Warning") | Out-Null; return }
        if ([string]::IsNullOrWhiteSpace($txtDomain.Text)) { [System.Windows.Forms.MessageBox]::Show($ui.EnterDomain, $ui.NoDomainTitle, "OK", "Warning") | Out-Null; return }

        $groups = @()
        if (-not [string]::IsNullOrWhiteSpace($txtGroups.Text)) {
            $groups = $txtGroups.Text.Split(',', [System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.Trim() }
        }

        Write-Log "$($ui.StartLogPrefix)$($txtOU.Text)$($ui.DomainPrefix)$($txtDomain.Text)$($ui.GroupsPrefix)$($groups -join ', ')" "INFO"

        $result = Create-UsersFromExcelData `
            -Users $script:LoadedUsers `
            -OU $txtOU.Text `
            -DomainSuffix $txtDomain.Text.Trim() `
            -GroupsToAdd $groups `
            -PasswordNeverExpires ([bool]$chkNever.Checked)

        try {
            $createdRows = @($result | Where-Object { $_.Status -eq 'OK' -and -not [string]::IsNullOrWhiteSpace($_.password) })
            if ($createdRows.Count -gt 0) {
                $pdfLog = Save-PasswordLogPdf -CreatedRows $createdRows -DomainSuffix $txtDomain.Text.Trim() -OU $txtOU.Text.Trim() -OutputDirectory $txtPdfFolder.Text
                if ($pdfLog) {
                    Write-Log "$($ui.PdfSaved) $($pdfLog.path)" "OK"
                }
            }
        } catch {
            Write-Log "$($ui.PdfError) $($_.Exception.Message)" "WARN"
        }
    })

    $null = $form.ShowDialog()
=======
function Show-MainForm {
    $texts = Get-MainFormTextMap
    $ui = New-MainFormUi -Texts $texts
    Register-MainFormEvents -Ui $ui -Texts $texts
    $null = $ui.Form.ShowDialog()
>>>>>>> Stashed changes
}
