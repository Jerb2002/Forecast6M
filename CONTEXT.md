# 6M cash flow forecast

A workbook, written by a notebook, that turns a pasted budget into a weekly cash flow per
country, and lets a short-term revenue paste override the near weeks.

## Language

### Time

**Closed month**:
A P&L month whose amount is final. Marked `Actual` on the P&L sheet, which `Budget input` reads;
arrives with the CFO's budget paste. Says nothing about when the cash landed.
_Avoid_: actual month, actuals (on its own), settled month

**Past week**:
A cash flow week before the first week of the short-term paste. Shown as nothing under an
override; shown as budget under Budget only.
_Avoid_: actual week, closed week, history

**Live week**:
A cash flow week from the first pasted week onward. The five pasted weeks, then the budget-driven
weeks after them.
_Avoid_: forecast week, open week

**Cash week**:
One column on a cash flow sheet: an ISO week, keyed `2026-W40`. A month on a CF sheet is a label
over whole weeks, not a calendar month.
_Avoid_: week of month, period

### The paste

**Short-term paste**:
Five weeks of expected customer receipts per country, in DKK, pasted on `Short-term input` each
Monday. Built by the short-term tool from open invoice items: gross of VAT and duty, dated when the
customer is expected to pay. Factored invoices are not in it.
_Avoid_: short-term forecast (that is the tool, not the figures), override numbers, bank cash

**Bank cash**:
What reaches our account in a week: the unfactored receipts plus the bank's legs on the factored
ones (advance, retention less fee). What a cash flow sheet shows, in two groups with a total each:
`Paid by the customer` and `From factoring`. Under an override the paste stands in for the first
group, whose total in a pasted week is the paste itself; the second keeps the budget's figures.
_Avoid_: receipts (that is the customer's side), revenue cash

**Budget paste**:
The CFO's own budget workbook copied onto the eight P&L sheets and `Budget input`, roughly
quarterly. Carries the closed months.
_Avoid_: budget input (that is one sheet), the budget sheets

**Monday paste**:
The weekly routine on the live workbook: type the Monday, paste the short-term paste, save. Done
in Excel only; the notebook is never run for it.
_Avoid_: weekly update, refresh, re-run

### The file

**Rebuild**:
A run of the notebook. It writes a fresh workbook; nothing typed into the old one survives unless
carried across by hand. The only kind of re-run there is.
_Avoid_: refresh, re-run, regenerate, update

**Human-owned sheet**:
A sheet whose contents are typed or pasted in Excel and carried across a rebuild by hand: the
eight P&L sheets, `Factoring`, `COGS delay`, `Cost timing`, `Short-term input`, the Rate,
Charged and Reclaimed columns of `VAT`, and the basis switch cell.
_Avoid_: input sheet (that is a tab group), configurable sheet

**Notebook-owned sheet**:
A sheet the notebook writes in full on every rebuild and nobody edits: the cash flow, dashboard,
check, contents and Python Data sheets. Formulas on them read the human-owned sheets.
_Avoid_: python sheet, generated sheet, output sheet

### The switch

**Basis**:
Which figures a cash flow sheet shows: Budget only, Short-term override, or Difference. One
dropdown on `CF OVERVIEW ALL`, the workbook name `Basis`; every CF sheet reads it and none has a
switch of its own.
_Avoid_: mode, view, scenario

**Overridden week**:
A cash week that is one of the five pasted weeks, on a basis other than Budget only. The three
cash-in rows carry the paste there; every other row carries the budget.
_Avoid_: pasted week (the paste is the input; the week is the column), short-term week

**Plain key**:
The week key the untouched rows look up: the real cash week when the column is live on the
current basis, and `-` (a key that matches nothing) when it is not. How the basis reaches
twenty-seven rows without an IF in any of them.
_Avoid_: masked key, switch key

**Unshown week**:
A cash week column with no figure on the current basis: a week before the first pasted week
under Short-term override, and any week but the pasted ones under Difference. Its cells hold
nought and are greyed and blank, header included; the note cell says why. Not the same as a
week with no cash, which reads `-`. A month with no shown week is greyed with them; a month
partly shown totals the weeks it shows.
_Avoid_: past week (a week after the paste window is unshown too, under Difference), empty week

**Budget-basis row**:
A hidden row on a CF sheet holding a line's figure on the budget basis, keyed on the real cash
week whatever `Basis` says: the three cash-in lines, every purchase line, and the cost, VAT and
fee figures built off them. What `Check` reads, so it ties whichever basis the sheets are
showing. The same lookup the visible row carries, not a read of the tables behind it.
_Avoid_: shadow row, backup row, pinned row
