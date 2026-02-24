function Get-ExcelSheetWithColumn {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$ColumnName # напр. "Вступник"
    )

    Import-Module ImportExcel -ErrorAction Stop

    $pkg = Open-ExcelPackage -Path $Path
    try {
        foreach ($ws in $pkg.Workbook.Worksheets) {
            if (-not $ws.Dimension) { continue }

            $endCol = $ws.Dimension.End.Column
            $headers = @()
            for ($c=1; $c -le $endCol; $c++) {
                $h = $ws.Cells[1,$c].Text
                $h = (($h -replace '\u00A0',' ') -as [string]).Trim()
                $headers += $h
            }
            $norm = $headers | ForEach-Object { (($_ -as [string]).Trim().ToLower()) }
            if ($norm -contains $ColumnName.Trim().ToLower()) {
                return $ws.Name
            }
        }
    }
    finally {
        Close-ExcelPackage $pkg
    }

    return $null
}

function Import-UsersFromExcelSmart {
    param([Parameter(Mandatory=$true)][string]$Path)

    Import-Module ImportExcel -ErrorAction Stop

    $sheet = Get-ExcelSheetWithColumn -Path $Path -ColumnName "Вступник"
    if (-not $sheet) { throw "Не знайдено лист з колонкою 'Вступник' (заголовок у 1-му рядку)." }

    $users = Import-Excel -Path $Path -WorksheetName $sheet
    return [pscustomobject]@{ Sheet = $sheet; Users = $users }
}
