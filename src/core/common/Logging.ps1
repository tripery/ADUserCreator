function Set-LogTarget {
    param([Parameter(Mandatory=$true)][System.Windows.Forms.TextBox]$TextBox)
    $script:LogTextBox = $TextBox
}

function Write-Log {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet("INFO","WARN","ERROR","OK")][string]$Level = "INFO"
    )

    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "[$ts][$Level] $Message`r`n"

    if ($script:LogTextBox) {
        $script:LogTextBox.AppendText($line)
    } else {
        Write-Host $line
    }
}
