param([string]$Path = "C:\Users\admin\Desktop\Python\6MForecast\forecast_dev.xlsx")
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($Path)
$wb.Names.Item("Basis").RefersToRange.Value2 = "Budget only"
$xl.CalculateFull()
function Row($sheet, $label) { foreach ($c in $sheet.Range("B1:B80").Cells) { if ($c.Text -eq $label) { return $c.Row } } }
$labels = "VAT on sales", "VAT on purchases", "VAT settlement"
foreach ($cc in "DNK","NOR","SWE","USA") {
  $tot = @{}; foreach ($l in $labels) { $tot[$l] = 0.0 }
  $last = ""
  foreach ($y in 2026,2027,2028) {
    $ws = $wb.Worksheets.Item("CF $cc $y")
    # year total column = last used column in row 6 with 'Year' header
    $ycol = $null
    foreach ($c in $ws.Range("E6:JZ6").Cells) { if ($c.Text -eq "Year") { $ycol = $c.Column } }
    foreach ($l in $labels) { $r = Row $ws $l; if ($r) { $tot[$l] += [double]$ws.Cells.Item($r, $ycol).Value2 } }
    # last week with any settlement cash
    $r = Row $ws "VAT settlement"
    for ($c = $ycol - 1; $c -ge 5; $c--) { $v = $ws.Cells.Item($r, $c).Value2; if ($v -and [math]::Abs($v) -gt 0.5) { $last = "$($ws.Cells.Item(7,$c).Text)"; break } }
  }
  "{0}: sales={1:n0} purchases={2:n0} settlement={3:n0} net={4:n0} last settlement week={5}" -f $cc, $tot["VAT on sales"], $tot["VAT on purchases"], $tot["VAT settlement"], ($tot["VAT on sales"] - $tot["VAT on purchases"] - $tot["VAT settlement"]), $last
}
$ws = $wb.Worksheets.Item("VAT cash")
"VAT cash used range: " + $ws.UsedRange.Address($false,$false)
$wb.Close($false); $xl.Quit()
