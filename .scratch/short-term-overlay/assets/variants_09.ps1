# Ticket 09 prototype: how a past week reads. Copies forecast_dev.xlsx, pastes five weeks, sets
# the basis, then clones "CF DNK 2026" into four variant sheets, one formatting rule each, and
# exports a picture of each around the paste boundary. Throwaway.
# Usage: powershell -File variants_09.ps1 [-Basis "Short-term override"|"Difference"]
param([string]$Basis = "Short-term override")
$root = "C:\Users\admin\Desktop\Python\6MForecast"
$src = "$root\forecast_dev.xlsx"
$assets = "$root\.scratch\short-term-overlay\assets"
$out = "$assets\09-variants.xlsx"
Copy-Item $src $out -Force

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($out)
$xl.Calculation = -4135   # manual

# --- paste five weeks, same figures as ticket 05
$st = $wb.Worksheets.Item("Short-term input")
$amounts = @(@(1000000,1100000,1200000,1300000,1400000), @(2000000,2100000,2200000,2300000,2400000),
             @(3000000,3100000,3200000,3300000,3400000), @(4000000,4100000,4200000,4300000,4400000))
for ($i = 0; $i -lt 4; $i++) { for ($j = 0; $j -lt 5; $j++) { $st.Cells.Item(5 + $i, 3 + $j).Value2 = $amounts[$i][$j] } }
$dash = $wb.Worksheets.Item("Dashboard CF")
$dash.Cells.Item(7, 3).Formula = $Basis
$xl.Calculate()
$first = $st.Range("K11").Text
"basis=$Basis  first paste week=$first"

$base = $wb.Worksheets.Item("CF DNK 2026")
function Row($sheet, $label) { foreach ($c in $sheet.Range("B1:B80").Cells) { if ($c.Text -eq $label) { return $c.Row } } }
$liveRow = Row $base "Live on this basis"; $overRow = Row $base "Overridden by the paste"; $multRow = Row $base "Budget figure survives"
$firstRow = Row $base "Cash in"; $netRow = Row $base "Net cash flow"; $lastRow = Row $base "Factoring fee"
$keyRow = 7; $headRow = 6
# week columns vs month-total columns, and the last column (Year)
$weekCols = @(); $totalCols = @(); $lastCol = 0; $firstCol = 0
foreach ($c in $base.Range("E6:IZ6").Cells) {
    if ($c.Text -eq "") { break }
    if ($c.Text -eq "Total") { $totalCols += $c.Column } elseif ($c.Text -eq "Year") { $lastCol = $c.Column } else { $weekCols += $c.Column }
    if ($base.Cells.Item($keyRow, $c.Column).Text -eq $first) { $firstCol = $c.Column }
}
"rows: live=$liveRow over=$overRow mult=$multRow body=$firstRow..$lastRow net=$netRow; weeks=$($weekCols.Count) totals=$($totalCols.Count) year col=$lastCol; first paste col=$firstCol"

$GREY = 0xD9E0E1     # GRID #e1e0d9 as BGR
$MUTED = 0x818789    # MUTED #898781 as BGR
$xlExpression = 2

function ColL($n) { $s = ""; while ($n -gt 0) { $m = ($n - 1) % 26; $s = [string][char](65 + [int]$m) + $s; $n = [int][math]::Floor(($n - 1) / 26) }; return $s }

function AddVariant($name, $rule, $fmt, $fill, $greyHeader, $banner) {
    # $rule: a formula template with {c} for the column letter, true when the column is NOT shown
    $base.Copy([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
    $ws = $wb.Worksheets.Item($wb.Worksheets.Count); $ws.Name = $name
    $ws.Tab.Color = 0x00C0FF
    $ws.Cells.Item(2, 2).Formula = "$name - " + [string]$ws.Cells.Item(2, 2).Text
    # week columns: one rule over the whole body block, relative to the top-left cell
    # one rule per month's run of week columns, so a total column is never inside a run
    foreach ($t in $totalCols) {
        $prev = ($totalCols | Where-Object { $_ -lt $t } | Measure-Object -Maximum).Maximum
        if (-not $prev) { $prev = $weekCols[0] - 1 }
        $body = $ws.Range($ws.Cells.Item($firstRow, $prev + 1), $ws.Cells.Item($lastRow, $t - 1))
        $fc = $body.FormatConditions.Add($xlExpression, [Type]::Missing, ($rule -replace "\{c\}", (ColL ($prev + 1))))
        if ($fmt) { $fc.NumberFormat = $fmt }
        if ($fill) { $fc.Interior.Color = $GREY }
    }
    # month totals: a whole month is "not shown" when none of its weeks is; a partly-past month
    # keeps its sum of live weeks (what SUM does today)
    foreach ($t in $totalCols) {
        $prev = ($totalCols | Where-Object { $_ -lt $t } | Measure-Object -Maximum).Maximum
        if (-not $prev) { $prev = $weekCols[0] - 1 }
        $tr = $ws.Range($ws.Cells.Item($firstRow, $t), $ws.Cells.Item($lastRow, $t))
        $parts = @(); for ($k = $prev + 1; $k -lt $t; $k++) { $parts += "(" + (($rule -replace "\{c\}", (ColL $k)) -replace "^=", "") + ")" }
        $fc2 = $tr.FormatConditions.Add($xlExpression, [Type]::Missing, "=AND(" + ($parts -join ",") + ")")
        if ($fmt) { $fc2.NumberFormat = $fmt }
        if ($fill) { $fc2.Interior.Color = $GREY }
    }
    if ($greyHeader) {
        foreach ($t in $totalCols) {
            $prev = ($totalCols | Where-Object { $_ -lt $t } | Measure-Object -Maximum).Maximum
            if (-not $prev) { $prev = $weekCols[0] - 1 }
            $hdr = $ws.Range($ws.Cells.Item($headRow, $prev + 1), $ws.Cells.Item($headRow, $t - 1))
            $fc3 = $hdr.FormatConditions.Add($xlExpression, [Type]::Missing, ($rule -replace "\{c\}", (ColL ($prev + 1))))
            $fc3.Font.Color = $MUTED; $fc3.Interior.Color = $GREY
            $parts = @(); for ($k = $prev + 1; $k -lt $t; $k++) { $parts += "(" + (($rule -replace "\{c\}", (ColL $k)) -replace "^=", "") + ")" }
            $fc4 = $ws.Cells.Item($headRow, $t).FormatConditions.Add($xlExpression, [Type]::Missing, "=AND(" + ($parts -join ",") + ")")
            $fc4.Font.Color = $MUTED; $fc4.Interior.Color = $GREY
        }
    }
    if ($banner) { $ws.Range("B3").Formula = $banner }
    # a running total under the net, week columns only; a separate yes/no from the variants
    $ws.Rows.Item($netRow + 1).Insert() | Out-Null
    $cum = $netRow + 1
    $ws.Cells.Item($cum, 2).Value2 = "Cumulative net"
    $ws.Cells.Item($cum, 2).Font.Italic = $true
    $prevL = $null
    foreach ($k in $weekCols) {
        $L = ColL $k
        if ($prevL) { $ws.Cells.Item($cum, $k).Formula = "=$prevL$cum+$L$netRow" } else { $ws.Cells.Item($cum, $k).Formula = "=$L$netRow" }
        $ws.Cells.Item($cum, $k).NumberFormat = '#,##0,;(#,##0,);"-"'; $ws.Cells.Item($cum, $k).Font.Italic = $true
        $prevL = $L
    }
    return $ws
}

# not-shown rules, per column letter {c}
$past   = '={c}$' + $liveRow + '=0'                                                   # a week before the paste
$unseen = '=AND({c}$' + $multRow + '=0,{c}$' + $overRow + '=0)'                        # no figure on this basis at all

$bannerC = '="Basis: "&Basis&IF(Basis="Budget only","","   |   weeks before "&FirstPasteWeek&" are not shown")'
$bannerD = '="Basis: "&Basis&IF(Basis="Budget only","",IF(Basis="Difference","   |   only the five pasted weeks are shown","   |   weeks before "&FirstPasteWeek&" are not shown"))'

$A = AddVariant "09-A empty" $past ';;;' $false $false $null
$B = AddVariant "09-B marker" $past ';;"past"' $false $false $null
$C = AddVariant "09-C grey + banner" $past ';;;' $true $true $bannerC
$D = AddVariant "09-D grey, incl outside window" $unseen ';;;' $true $true $bannerD
$xl.Calculate()

# --- pictures: hide the week columns far from the boundary, snapshot, unhide
$lo = $firstCol - 5; $hi = $firstCol + 8
foreach ($ws in $A, $B, $C, $D) {
    foreach ($k in $weekCols + $totalCols) { if ($k -lt $lo -or $k -gt $hi) { $ws.Columns.Item($k).Hidden = $true } }
    $ws.Columns.Item($lastCol).Hidden = $true
    $rng = $ws.Range($ws.Cells.Item(2, 2), $ws.Cells.Item($lastRow + 1, $lastCol))
    $rng.CopyPicture(1, 2) | Out-Null
    $co = $ws.ChartObjects().Add(0, 0, $rng.Width, $rng.Height)
    $co.Chart.Paste() | Out-Null
    $png = "$assets\09-" + ($ws.Name -replace '[^A-Za-z0-9]+', '-') + "-" + ($Basis -replace '[^A-Za-z]+', '-') + ".png"
    $co.Chart.Export($png) | Out-Null
    $co.Delete()
    foreach ($k in $weekCols + $totalCols) { $ws.Columns.Item($k).Hidden = $false }
    $ws.Columns.Item($lastCol).Hidden = $false
    "wrote $png"
}

# error cells on the variants
foreach ($ws in $A, $B, $C, $D) {
    $n = 0; try { $r = $ws.UsedRange.SpecialCells(-4123, 16); if ($r) { $n = $r.Count } } catch {}
    "$($ws.Name): error cells $n; cumulative at last week = " + $ws.Cells.Item($netRow + 1, $weekCols[-1]).Text
}
$xl.Calculation = -4105
$wb.Save(); $wb.Close($true); $xl.Quit()
"saved $out"


