# The three-way basis dropdown

Type: prototype
Status: resolved
Claimed by: claude-opus-5 session 263fd5e4
Resolved at: 2026-09-14T22:10:00+02:00
Blocked by: 02, 03
Assets: .scratch/short-term-overlay/assets/patch_basis.py, drive_basis.ps1, time_basis.ps1

## Question

The CF sheets get a dropdown with three settings: **Budget only**, **Short-term override**, and
**Difference**. Budget only is what the sheet does today. Override swaps the near five weeks for
the pasted numbers. Difference shows the gap, so the quality of the short-term prediction is a
number to read rather than something to eyeball by flipping.

Settle:

- **Where the switch lives.** One cell on `Dashboard CF` that all twelve CF sheets read, or a
  switch per sheet. One cell is one place to look and one thing to forget you set. Per sheet lets
  two countries be compared side by side.
- **How a switch reaches the formulas.** Every figure on a CF sheet is a `SUMIFS` on
  `country|line|week`. There are roughly fifteen thousand of them per sheet. Wrapping each in an
  `IF` is the obvious move and the expensive one. The alternative is making the key itself carry
  the basis, so the switch changes what is looked up rather than whether a lookup happens.
- **What Difference means above a total row.** The gap on `Revenue` is clear. The gap on
  `Cash out total`, where nothing is overridden, is nought - useful, or noise?
- **What Difference means outside the five weeks,** where there is nothing to compare.

Build it, open it in Excel, and check what it does to recalculation time before settling. A sheet
that takes ten seconds to switch is a sheet nobody switches.

**Added by ticket 03.** An overridden week is not one lookup on `ShortTerm`. Per country-week:
split the paste by the budget's duty ratio and VAT rate (`service = paste / (1 + r + v)`), then
add the budget's non-direct legs (`advance`, `retention`) from the payment table on `Revenue`,
`Revenue duty` and `VAT on sales`. That is a by-leg `SUMIFS` with three criteria on rows that use
one today, and it is on three rows per sheet, not thirty. Measure it. The memo fee row is
unchanged under override.

**Added by ticket 04.** Mondays are Excel-only, so every formula the switch needs is in the written
workbook; nothing can be computed in the notebook on a Monday number. And every re-run is a
rebuild, so wherever the switch cell lands it is human-owned: add it to the rebuild checklist on
the README, and prefer a place a person will remember to carry across (the `Dashboard CF` option
is one cell; per-sheet is twelve).

## Answer

**One dropdown on `Dashboard CF` (C7), named `Basis`. The switch reaches the sheets by changing
what is looked up, not whether. Built, driven in Excel 16 over COM, and timed against the
IF-wrapped alternative: no difference, and no cost over the workbook's own recalculation
floor.**

**Where it lives.** C7 on `Dashboard CF`, under Country and Compare with, the same blue entry
style and the same protection exception as the two dropdowns above it. The workbook name
`Basis` points at it and every CF sheet reads the name. Each CF sheet says `Basis: ...` in the
note cell under its title; that cell is a label. A per-sheet cell that could be typed over to
pin one sheet to another basis was built and shown, and the user had it removed: nobody asked
for it. The dashboard caption also names the basis, and the dashboard follows the sheets it
reads. C7 is human-owned and is on the README rebuild checklist; a fresh build opens on
*Budget only*.

**How it reaches the formulas.** Thirteen hidden rows under the key row on every CF sheet, one
figure per week column: `live` (Budget only, or the week is on or after `FirstPasteWeek`, a
second name on the first week key of `Short-term input`), `override` (not Budget only, and
the week is in `ShortTerm`), `mult` (live and not Difference), a *plain key* that is the real
week when `mult` is 1 and `-` otherwise, the paste, the month's budget duty ratio, the service
part of the paste, and for each of the three replaced rows the budget figure and its
`Direct` leg. The twenty-seven untouched rows look their week up off the plain key and are
otherwise the one SUMIFS they always were; `VAT on purchases` and the totals follow from the
rows above them. Only `Revenue`, `Revenue duty` and `VAT on sales` carry an IF: in an
overridden week, the paste's part plus the budget's factored legs (budget less direct leg),
less the whole budget figure again under Difference; in any other week, the budget figure
times `mult`.

**Measured** (Excel 16, manual calculation, one Calculate per step, both variants):

| step | plain key | IF on every cell |
|---|---|---|
| full recalculation | 2.7s | 2.8s |
| Calculate with no edit (the volatile floor: `TODAY()`, `INDIRECT`) | 1.7s | 1.7s |
| paste 20 figures | 2.1s | 2.2s |
| switch basis, any direction | 1.7-1.8s | 1.6-1.8s |
| edit one pasted figure | 1.6s | 1.7s |
| move the Monday a week | 1.7s | 1.7s |

Everything sits on the floor; the ticket's fear of a ten-second switch does not arise either
way. The plain-key version is kept because it leaves the untouched rows untouched; the IF
version's code was deleted from the notebook. A full recalculation was 3.2s before this
ticket: the country VAT rate is now one hidden cell per sheet rather than a lookup inlined
into every VAT cell.

**Checked in Excel.** Zero error cells on every basis. DNK 2026-W38 with 1,000,000 pasted:
service 676,030, and the three rows add back to the paste plus the budget's factored legs;
`Revenue` 1,382,238 under override against 1,326,012 budget, `Difference` 56,226. USA (no
VAT, no factoring) and NOR (no factoring): `Cash in total` in an overridden week equals the
paste exactly. Past weeks are nought on every row, and a month whose weeks are partly past
totals its live weeks. `Check` reads *OK* on all four countries under Budget only and *Off*
under the other two.

**Difference above a total row** (asked, answered: keep). Totals stay sums of the rows above,
so under Difference every cost row is nought, `Cash out total` is `-`, and `Net cash flow` is
the three revenue gaps added up. A plus is the short-term tool seeing more than the budget.

**Difference outside the five weeks** (asked, handed on). Both the weeks after the window and
the weeks before it come to nought and read `-`, indistinguishable from each other and from
a week with no cash. Cells hold nought rather than nothing so the `+`/`-` total formulas keep
working. How a past or an out-of-window week *reads* is ticket 09, which now covers both.

**Words** added to `CONTEXT.md`: *overridden week*, *plain key*; *basis* expanded.

**Facts handed on.** Ticket 07: `Check` reads the visible CF rows, so it says *Off* the moment
the basis leaves Budget only; the budget-basis figures for the three replaced rows survive
in the hidden `*_budget` rows, the cost rows do not. Ticket 09: cells hold 0 not blank; the
`live` row is the per-column flag a format or banner can key off; month subtotals currently
sum the live weeks.

