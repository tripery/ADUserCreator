function Show-MainForm {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $scriptDir = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($scriptDir)) {
        if (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
            $scriptDir = Split-Path -Parent $PSCommandPath
        }
        elseif ($MyInvocation -and $MyInvocation.MyCommand -and -not [string]::IsNullOrWhiteSpace($MyInvocation.MyCommand.Path)) {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
        }
        else {
            throw 'Unable to resolve ui script directory.'
        }
    }

    . (Join-Path $scriptDir 'Form.Localization.ps1')
    . (Join-Path $scriptDir 'Form.Layout.ps1')
    . (Join-Path $scriptDir 'Form.Events.ps1')

    $texts = Get-MainFormTexts
    $ui = New-MainFormLayout -Texts $texts
    Register-MainFormEvents -Ui $ui -Texts $texts

    [void]$ui.Form.ShowDialog()
}
