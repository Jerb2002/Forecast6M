# Ticket 10: open the rebuilt workbook in Excel 16, paste five weeks, walk the three bases,
# and read the whole Check sheet back under each: the verdict row, the three difference rows,
# the VAT block, the note cell. Then a picture of the sheet under Short-term override.
# Usage: powershell -File drive_check.ps1 [<xlsx>]
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
"first paste week " + $st.Range("K11").Text
"check note: '" + $check.Range("B3").Text + "'"

function Row($sheet, $label) { foreach ($c in $sheet.Range("B1:B80").Cells) { if ($c.Text -eq $label) { return $c.Row } } }
$errors = 0; foreach ($ws in $wb.Worksheets) { try { $r = $ws.UsedRange.SpecialCells(-4123, 16); if ($r) { $errors += $r.Count } } catch {} }
"error cells: $errors"

# the whole Check sheet, as text, row by row
function Dump($sheet) {
    $last = $sheet.UsedRange.Rows.Count + 1
    for ($r = 4; $r -le $last; $r++) {
        $label = $sheet.Cells.Item($r, 2).Text
        if ($label -eq "") { continue }
        $vals = @(); for ($c = 3; $c -le 7; $c++) { $vals += ("{0,18}" -f $sheet.Cells.Item($r, $c).Text) }
        ("  {0,-44}{1}" -f $label, ($vals -join ""))
    }
}

# the hidden budget-basis rows on CF DNK 2026, year column, against the visible rows on Budget only
$cf = $wb.Worksheets.Item("CF DNK 2026")
$ycol = $null; foreach ($c in $cf.Range("E6:JZ6").Cells) { if ($c.Text -eq "Year") { $ycol = $c.Column } }
"CF DNK 2026 year column $ycol; rows: " + (Row $cf "Cost cash, budget") + " cost_budget, " + (Row $cf "VAT on purchases, budget") + " vat_out_budget, " + (Row $cf "Revenue") + " Revenue (first visible)"

foreach ($basis in "Budget only", "Short-term override", "Difference") {
    $dash.Cells.Item(7, 3).Formula = $basis
    $t0 = Get-Date; $xl.Calculate(); $calc = ((Get-Date) - $t0).TotalSeconds
    ""; ("=== {0}  calc={1:n2}s" -f $basis, $calc)
    Dump $check
    # hidden vs visible on DNK 2026, year column: the copies must equal the visible rows on
    # Budget only and stay put on the other two
    $pairs = @(@("Cost cash, budget", "Cash out total"), @("VAT on purchases, budget", "VAT on purchases"), @("VAT settlement, budget", "VAT settlement"), @("Factoring fee, budget", "Factoring fee"), @("VAT on sales, budget", "VAT on sales"))
    foreach ($p in $pairs) {
        $h = $cf.Cells.Item((Row $cf $p[0]), $ycol).Value2; $v = $cf.Cells.Item((Row $cf $p[1]), $ycol).Value2
        ("  DNK 2026 year: {0,-28} hidden={1,14:n0}  visible {2,-18}={3,14:n0}" -f $p[0], $h, $p[1], $v)
    }
}

# picture under Override
$dash.Cells.Item(7, 3).Formula = "Short-term override"; $xl.Calculate()
$check.Activate()
$lastRow = $check.UsedRange.Rows.Count + 1
$rng = $check.Range($check.Cells.Item(2, 2), $check.Cells.Item($lastRow, 7))
$rng.CopyPicture(1, 2) | Out-Null
$co = $check.ChartObjects().Add(10, 10, $rng.Width + 4, $rng.Height + 4); $co.Activate() | Out-Null; $co.Chart.Paste() | Out-Null
$png = "$assets\10-built-Check-Short-term-override.png"; $co.Chart.Export($png) | Out-Null; $co.Delete()
"wrote $png"
$wb.Close($false); $xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
