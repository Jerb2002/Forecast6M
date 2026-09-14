# Ticket 05: open the built workbook in Excel 16, paste five weeks, walk the three bases, time
# each switch, and read the overridden rows back. Usage: powershell -File drive_basis.ps1 <xlsx>
param([string]$Path = "C:\Users\admin\Desktop\Python\6MForecast\forecast_dev.xlsx")

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$t0 = Get-Date; $wb = $xl.Workbooks.Open($Path); $open = ((Get-Date) - $t0).TotalSeconds
$t0 = Get-Date; $xl.CalculateFull(); $full1 = ((Get-Date) - $t0).TotalSeconds
$t0 = Get-Date; $xl.CalculateFull(); $full2 = ((Get-Date) - $t0).TotalSeconds
"file: $Path"
"open={0:n1}s  fullcalc={1:n1}s / {2:n1}s" -f $open, $full1, $full2

# --- error cells anywhere
$errors = 0; $where = @()
foreach ($ws in $wb.Worksheets) {
    try {
        $r = $ws.UsedRange.SpecialCells(-4123, 16)   # xlCellTypeFormulas, xlErrors
        if ($r) { $errors += $r.Count; $where += "$($ws.Name):$($r.Address($false,$false))" }
    } catch {}
}
"error cells: $errors $($where -join ' ')"

# --- names
foreach ($n in "Basis", "FirstPasteWeek") {
    $nm = $wb.Names.Item($n); "name $n -> $($nm.RefersTo) = $($nm.RefersToRange.Text)"
}

# --- paste five weeks: distinct round figures per country so a wrong join shows
$st = $wb.Worksheets.Item("Short-term input")
$amounts = @(@(1000000,1100000,1200000,1300000,1400000),
             @(2000000,2100000,2200000,2300000,2400000),
             @(3000000,3100000,3200000,3300000,3400000),
             @(4000000,4100000,4200000,4300000,4400000))
for ($i = 0; $i -lt 4; $i++) { for ($j = 0; $j -lt 5; $j++) {
    $st.Cells.Item(5 + $i, 3 + $j).Value2 = $amounts[$i][$j] } }
"pasted; countries: " + (($st.Range("B5:B8").Value2 | ForEach-Object { $_ }) -join ",")
"window: " + (($st.Range("C4:G4").Value2 | ForEach-Object { $_ }) -join " ")

$dash = $wb.Worksheets.Item("Dashboard CF")
$cf = $wb.Worksheets.Item("CF DNK 2026")
$check = $wb.Worksheets.Item("Check")

# the columns of the five pasted weeks on CF DNK 2026, and the two before them
$first = $st.Range("K11").Text
$keys = $cf.Range("E7:IZ7")
$col = $null
foreach ($c in $keys.Cells) { if ($c.Text -eq $first) { $col = $c.Column; break } }
"first pasted week $first is column $col on CF DNK 2026"
$cols = ($col - 2)..($col + 5)

function Row($sheet, $label) {
    foreach ($c in $sheet.Range("B1:B80").Cells) { if ($c.Text -eq $label) { return $c.Row } }
}
$rows = @{}
foreach ($l in "Revenue", "Revenue duty", "VAT on sales", "Cash in total", "COGS", "Cash out total", "Net cash flow", "Factoring fee", "Live on this basis", "Overridden by the paste", "Pasted revenue cash", "Budget duty over revenue", "Service part of the paste", "Revenue, budget", "Revenue, direct leg only", "VAT on sales, budget", "VAT on sales, direct leg only") {
    $rows[$l] = Row $cf $l
}
"rows: " + (($rows.GetEnumerator() | Sort-Object Value | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join " ")

foreach ($basis in "Budget only", "Short-term override", "Difference") {
    $dash.Cells.Item(7, 3).Formula = [string]$basis
    $t0 = Get-Date; $xl.Calculate(); $calc = ((Get-Date) - $t0).TotalSeconds
    ""
    ("=== {0}   calc={1:n2}s   CF B3='{2}' C3={3} D3={4}   dash caption='{5}'" -f $basis, $calc, $cf.Range("B3").Text, $cf.Range("C3").Value2, $cf.Range("D3").Value2, $dash.Range("B9").Text)
    "check verdict: " + (($check.Range("C" + (Row $check "Does it tie?") + ":G" + (Row $check "Does it tie?")).Value2 | ForEach-Object { $_ }) -join " | ")
    $head = "{0,-28}" -f "week"
    foreach ($c in $cols) { $head += "{0,12}" -f $cf.Cells.Item(7, $c).Text }
    $head
    foreach ($l in "Live on this basis", "Overridden by the paste", "Pasted revenue cash", "Budget duty over revenue", "Service part of the paste", "Revenue, budget", "Revenue, direct leg only", "Revenue", "Revenue duty", "VAT on sales, budget", "VAT on sales, direct leg only", "VAT on sales", "Cash in total", "COGS", "Cash out total", "Net cash flow", "Factoring fee") {
        $line = "{0,-28}" -f $l
        foreach ($c in $cols) {
            $v = $cf.Cells.Item($rows[$l], $c).Value2
            if ($v -is [double]) { $line += "{0,12:n0}" -f $v } else { $line += "{0,12}" -f $v }
        }
        $line
    }
}

# a second switch back, to time it warm
$dash.Cells.Item(7, 3).Formula = "Short-term override"
$t0 = Get-Date; $xl.Calculate(); $calc = ((Get-Date) - $t0).TotalSeconds
""; "switch back to override: {0:n2}s" -f $calc
# a Monday paste edit: one figure changes
$t0 = Get-Date; $st.Cells.Item(5, 3).Value2 = 1500000; $xl.Calculate(); $calc = ((Get-Date) - $t0).TotalSeconds
"edit one pasted figure: {0:n2}s  -> DNK Revenue week1 now {1:n0}" -f $calc, $cf.Cells.Item($rows["Revenue"], $col).Value2

$wb.Close($false); $xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
