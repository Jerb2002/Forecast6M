# Ticket 09: open the rebuilt workbook in Excel 16, paste five weeks, walk the three bases,
# and read back how an unshown week displays (fill, text) on every CF sheet's first past
# week, plus the note cell and the Check verdict. Pictures of CF DNK 2026 under each basis.
# Usage: powershell -File drive_shown.ps1 [<xlsx>]
param([string]$Path = "C:\Users\admin\Desktop\Python\6MForecast\forecast_dev.xlsx")
$assets = "C:\Users\admin\Desktop\Python\6MForecast\.scratch\short-term-overlay\assets"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($Path)
$xl.Calculation = -4135
$t0 = Get-Date; $xl.CalculateFull(); "fullcalc={0:n1}s" -f ((Get-Date) - $t0).TotalSeconds

$st = $wb.Worksheets.Item("Short-term input")
$amounts = @(@(1000000,1100000,1200000,1300000,1400000), @(2000000,2100000,2200000,2300000,2400000),
             @(3000000,3100000,3200000,3300000,3400000), @(4000000,4100000,4200000,4300000,4400000))
for ($i = 0; $i -lt 4; $i++) { for ($j = 0; $j -lt 5; $j++) { $st.Cells.Item(5 + $i, 3 + $j).Value2 = $amounts[$i][$j] } }
$dash = $wb.Worksheets.Item("Dashboard CF"); $check = $wb.Worksheets.Item("Check")
$xl.Calculate()
$first = $st.Range("K11").Text; "first paste week $first"

function Row($sheet, $label) { foreach ($c in $sheet.Range("B1:B80").Cells) { if ($c.Text -eq $label) { return $c.Row } } }
$cf = $wb.Worksheets.Item("CF DNK 2026")
$shownRow = Row $cf "Figure on this basis"; $revRow = Row $cf "Revenue"; $netRow = Row $cf "Net cash flow"; $lastRow = Row $cf "Factoring fee"
$weekCols = @(); $totalCols = @(); $lastCol = 0; $firstCol = 0
foreach ($c in $cf.Range("E6:IZ6").Cells) {
    if ($c.Text -eq "") { break }
    if ($c.Text -eq "Total") { $totalCols += $c.Column } elseif ($c.Text -eq "Year") { $lastCol = $c.Column } else { $weekCols += $c.Column }
    if ($cf.Cells.Item(7, $c.Column).Text -eq $first) { $firstCol = $c.Column }
}
"shown row $shownRow, revenue row $revRow, first paste col $firstCol, year col $lastCol"

$errors = 0; foreach ($ws in $wb.Worksheets) { try { $r = $ws.UsedRange.SpecialCells(-4123, 16); if ($r) { $errors += $r.Count } } catch {} }
"error cells: $errors"

foreach ($basis in "Budget only", "Short-term override", "Difference") {
    $dash.Cells.Item(7, 3).Formula = $basis
    $t0 = Get-Date; $xl.Calculate(); $calc = ((Get-Date) - $t0).TotalSeconds
    ""; ("=== {0}  calc={1:n2}s  note='{2}'" -f $basis, $calc, $cf.Range("B3").Text)
    "check: " + (($check.Range("C" + (Row $check "Does it tie?") + ":G" + (Row $check "Does it tie?")).Value2 | ForEach-Object { $_ }) -join " | ")
    # the week before the paste, the first pasted week, the week after the window, and the
    # month totals either side: shown flag, revenue text, and the fill Excel displays
    foreach ($col in ($firstCol - 1), $firstCol, ($firstCol + 5), ($totalCols | Where-Object { $_ -lt $firstCol } | Select-Object -Last 1), ($totalCols | Where-Object { $_ -gt $firstCol } | Select-Object -First 1), $lastCol) {
        $cell = $cf.Cells.Item($revRow, $col); $head = $cf.Cells.Item(6, $col)
        ("  col {0,-6} shown={1,-3} head='{2}' headfill={3,-8} rev text='{4}' value={5,12:n0} fill={6}" -f $head.Text, $cf.Cells.Item($shownRow, $col).Text, $head.Text, $head.DisplayFormat.Interior.Color, $cell.Text, $cell.Value2, $cell.DisplayFormat.Interior.Color)
    }
    # every CF sheet: does the week before the paste grey out (2026 sheets), and 2027 never
    foreach ($ws in $wb.Worksheets) {
        if ($ws.Name -notlike "CF *") { continue }
        $c = $ws.Cells.Item($revRow, $firstCol - 1)
        ("  {0,-12} col before paste: text='{1}' fill={2}" -f $ws.Name, $c.Text, $c.DisplayFormat.Interior.Color)
    }
    # picture
    $cf.Activate()
    foreach ($k in $weekCols + $totalCols) { $cf.Columns.Item($k).Hidden = ($k -lt $firstCol - 6 -or $k -gt $firstCol + 9) }
    $cf.Columns.Item($lastCol).Hidden = $true
    $rng = $cf.Range($cf.Cells.Item(2, 2), $cf.Cells.Item($lastRow, $lastCol))
    $rng.CopyPicture(1, 2) | Out-Null
    $co = $cf.ChartObjects().Add(10, 10, $rng.Width + 4, $rng.Height + 4); $co.Activate() | Out-Null; $co.Chart.Paste() | Out-Null
    $png = "$assets\09-built-" + ($basis -replace '[^A-Za-z]+', '-') + ".png"; $co.Chart.Export($png) | Out-Null; $co.Delete()
    foreach ($k in $weekCols + $totalCols) { $cf.Columns.Item($k).Hidden = $false }; $cf.Columns.Item($lastCol).Hidden = $false
    "  wrote $png"
}
$wb.Close($false); $xl.Quit()
