# Dashboard CF by week: open the rebuilt workbook in Excel 16 and read the dashboard back
# against the CF sheet it points at, for two years; picture of the block.
# Usage: powershell -File drive_dash_weeks.ps1 [<xlsx>]
param([string]$Path = "C:\Users\admin\Desktop\Python\6MForecast\forecast_dev.xlsx")
$assets = "C:\Users\admin\Desktop\Python\6MForecast\.scratch\short-term-overlay\assets"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($Path)
$xl.Calculation = -4135
$t0 = Get-Date; $xl.CalculateFull(); "fullcalc={0:n1}s" -f ((Get-Date) - $t0).TotalSeconds
$errors = 0; foreach ($ws in $wb.Worksheets) { try { $r = $ws.UsedRange.SpecialCells(-4123, 16); if ($r) { $errors += $r.Count } } catch {} }
"error cells: $errors"

function Row($sheet, $label) { foreach ($c in $sheet.Range("B1:B80").Cells) { if ($c.Text -eq $label) { return $c.Row } } }
$dash = $wb.Worksheets.Item("Dashboard CF")
$dash.Unprotect()
"caption: " + $dash.Range("B9").Text

foreach ($year in 2026, 2027) {
    $dash.Cells.Item(5, 5).Formula = "$year"; $dash.Cells.Item(5, 3).Formula = "DNK"; $xl.Calculate()
    $cf = $wb.Worksheets.Item("CF DNK $year")
    ""; "=== DNK $year"
    # month row and week headers, first 12 columns and the last three
    $months = @(); $weeks = @(); for ($c = 3; $c -le 55; $c++) { $months += $dash.Cells.Item(10, $c).Text; $weeks += $dash.Cells.Item(11, $c).Text }
    "months: " + (($months | Where-Object { $_ -ne "" }) -join " ")
    "weeks 1-6: " + ($weeks[0..5] -join " ") + "  ... 51-53: '" + ($weeks[50..52] -join "' '") + "'"
    # every week of Revenue and Net cash flow on the dashboard against the CF sheet
    foreach ($label in "Revenue", "Net cash flow") {
        $dr = Row $dash $label; $cr = Row $cf $label
        $mismatch = 0; $n = 0
        for ($c = 3; $c -le 55; $c++) {
            $t = $dash.Cells.Item($dr, $c).Text; if ($t -eq "") { continue }
            $n++
            $key = "$year-W" + ("{0:d2}" -f ($c - 2))
            $cfcol = $null; foreach ($k in $cf.Range("E7:JZ7").Cells) { if ($k.Text -eq $key) { $cfcol = $k.Column; break } }
            $a = [double]$dash.Cells.Item($dr, $c).Value2; $b = [double]$cf.Cells.Item($cr, $cfcol).Value2
            if ([math]::Abs($a - $b) -gt 0.5) { $mismatch++; "  MISMATCH $label $key dash=$a cf=$b" }
        }
        $ycol = $null; foreach ($k in $cf.Range("E6:JZ6").Cells) { if ($k.Text -eq "Year") { $ycol = $k.Column } }
        $tot = [double]$dash.Cells.Item($dr, 57).Value2; $cfy = [double]$cf.Cells.Item($cr, $ycol).Value2
        ("  {0,-14} {1} weeks shown, {2} mismatches; dashboard total {3:n0} vs CF year {4:n0}" -f $label, $n, $mismatch, $tot, $cfy)
    }
    "  compare header: '" + $dash.Cells.Item(11, 59).Text + "'  variance hdr: '" + $dash.Cells.Item(11, 60).Text + "'"
}

# picture: DNK 2026, the block, first ~16 week columns
$dash.Cells.Item(5, 5).Formula = "2026"; $xl.Calculate()
$dash.Activate()
for ($c = 19; $c -le 55; $c++) { $dash.Columns.Item($c).Hidden = $true }
$last = Row $dash "Factoring fee"
$rng = $dash.Range($dash.Cells.Item(2, 2), $dash.Cells.Item($last, 61))
$rng.CopyPicture(1, 2) | Out-Null
$co = $dash.ChartObjects().Add(10, 10, $rng.Width + 4, $rng.Height + 4); $co.Activate() | Out-Null; $co.Chart.Paste() | Out-Null
$png = "$assets\dash-weeks-DNK-2026.png"; $co.Chart.Export($png) | Out-Null; $co.Delete()
for ($c = 19; $c -le 55; $c++) { $dash.Columns.Item($c).Hidden = $false }
"wrote $png"
$wb.Close($false); $xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
