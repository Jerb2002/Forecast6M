# Ticket 09: picture each variant sheet of 09-variants.xlsx around the paste boundary, under
# one basis. Excel has to be visible for CopyPicture/Paste to carry a bitmap. Throwaway.
param([string]$Basis = "Short-term override")
$assets = "C:\Users\admin\Desktop\Python\6MForecast\.scratch\short-term-overlay\assets"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open("$assets\09-variants.xlsx")
$wb.Worksheets.Item("Dashboard CF").Cells.Item(7, 3).Formula = $Basis
$xl.Calculate()
$first = $wb.Worksheets.Item("Short-term input").Range("K11").Text
$base = $wb.Worksheets.Item("CF DNK 2026")
$weekCols = @(); $totalCols = @(); $lastCol = 0; $firstCol = 0
foreach ($c in $base.Range("E6:IZ6").Cells) {
    if ($c.Text -eq "") { break }
    if ($c.Text -eq "Total") { $totalCols += $c.Column } elseif ($c.Text -eq "Year") { $lastCol = $c.Column } else { $weekCols += $c.Column }
    if ($base.Cells.Item(7, $c.Column).Text -eq $first) { $firstCol = $c.Column }
}
$lo = $firstCol - 6; $hi = $firstCol + 9
foreach ($ws in $wb.Worksheets) {
    if ($ws.Name -notlike "09-*") { continue }
    $ws.Activate()
    foreach ($k in $weekCols + $totalCols) { $ws.Columns.Item($k).Hidden = ($k -lt $lo -or $k -gt $hi) }
    $ws.Columns.Item($lastCol).Hidden = $true
    $rng = $ws.Range($ws.Cells.Item(2, 2), $ws.Cells.Item(59, $lastCol))
    $rng.CopyPicture(1, 2) | Out-Null
    $co = $ws.ChartObjects().Add(10, 10, $rng.Width + 4, $rng.Height + 4)
    $co.Activate() | Out-Null
    $co.Chart.Paste() | Out-Null
    $png = "$assets\09-" + ($ws.Name -replace '[^A-Za-z0-9]+', '-') + "-" + ($Basis -replace '[^A-Za-z]+', '-') + ".png"
    $co.Chart.Export($png) | Out-Null
    $co.Delete()
    foreach ($k in $weekCols + $totalCols) { $ws.Columns.Item($k).Hidden = $false }
    $ws.Columns.Item($lastCol).Hidden = $false
    "wrote $png"
}
$wb.Close($false); $xl.Quit()
