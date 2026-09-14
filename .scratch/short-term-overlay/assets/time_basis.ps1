# Ticket 05: the recalculation clock, on one workbook. Same sequence for each variant so the
# numbers compare. Usage: powershell -File time_basis.ps1 <xlsx>
param([string]$Path)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
$wb = $xl.Workbooks.Open($Path)
$xl.Calculation = -4135   # manual: one Calculate per line below
$xl.CalculateFull() | Out-Null
function Clock($label, $block) {
    $t0 = Get-Date; & $block; $xl.Calculate(); $s = ((Get-Date) - $t0).TotalSeconds
    "{0,-42} {1,6:n2}s" -f $label, $s
}
$st = $wb.Worksheets.Item("Short-term input")
$dash = $wb.Worksheets.Item("Dashboard CF")
$ct = $wb.Worksheets.Item("Cost timing")
$cf = $wb.Worksheets.Item("CF DNK 2026")
"file: " + (Split-Path $Path -Leaf)
"{0,-42} {1,6:n2}s" -f "full recalculation", ((Measure-Command { $xl.CalculateFull() }).TotalSeconds)
"{0,-42} {1,6:n2}s" -f "full recalculation, again", ((Measure-Command { $xl.CalculateFull() }).TotalSeconds)
Clock "no edit, Calculate (volatile floor)" {}
Clock "paste 20 figures, Budget only" { for ($i = 0; $i -lt 4; $i++) { for ($j = 0; $j -lt 5; $j++) { $st.Cells.Item(5 + $i, 3 + $j).Value2 = 1000000 * ($i + 1) + 100000 * $j } } }
Clock "switch Budget only -> Override" { $dash.Cells.Item(7, 3).Formula = "Short-term override" }
Clock "edit one pasted figure, Override" { $st.Cells.Item(5, 3).Value2 = 1500000 }
Clock "edit one pasted figure again, Override" { $st.Cells.Item(5, 3).Value2 = 1600000 }
Clock "switch Override -> Difference" { $dash.Cells.Item(7, 3).Formula = "Difference" }
Clock "edit one pasted figure, Difference" { $st.Cells.Item(5, 4).Value2 = 1700000 }
Clock "switch Difference -> Budget only" { $dash.Cells.Item(7, 3).Formula = "Budget only" }
Clock "move the Monday a week, Budget only" { $st.Cells.Item(5, 10).Formula = "=DATE(2026,9,21)" }
Clock "switch Budget only -> Override" { $dash.Cells.Item(7, 3).Formula = "Short-term override" }
Clock "move the Monday a week, Override" { $st.Cells.Item(5, 10).Formula = "=DATE(2026,9,28)" }
Clock "pin CF DNK 2026 to Difference (B3)" { $cf.Cells.Item(3, 2).Formula = "Difference" }
"CF DNK 2026 B3 = " + $cf.Cells.Item(3, 2).Text + "; CF NOR 2026 B3 = " + $wb.Worksheets.Item("CF NOR 2026").Cells.Item(3, 2).Text
$wb.Close($false); $xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
