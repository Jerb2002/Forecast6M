# Map: the short-term revenue overlay

## Destination

`forecast_dev.ipynb` writes a workbook whose cash flow sheets show, per country and week:
closed-past on an actuals basis, the pasted 5-week short-term revenue forecast for the near
weeks, and the existing budget-driven long-term model for every week after that - switchable
on the sheet between *Budget only*, *Short-term override* and *Difference*, and re-pasted every
Monday without re-running the notebook.

Reached when a CFO can open the workbook on a Monday, paste five weeks of revenue cash per
country, and read a cash flow that uses it - with the Check sheet still tying CF to Budget on
the budget basis alone.

## Notes

**This map executes.** Work is carried into `forecast_dev.ipynb`, not handed off as a spec.
`forecast_final.ipynb` and `forecast_final.xlsx` are not touched by this effort.

- Build one notebook section at a time, and stop between steps. Do not add checks nobody asked for.
- Where a layout is in question, build the options as lettered variants and compare them in the
  real workbook. Do not settle a layout by argument.
- Excel 16 is on this machine. Drive it from PowerShell COM to open the written workbook and
  check it, rather than trusting openpyxl's view of what was written.
- Skills to call: `mattpocock-skills:grilling` and `mattpocock-skills:domain-modeling` on every
  grilling ticket; `mattpocock-skills:prototype` on every prototype ticket.

### Established before charting, not decisions

- `forecast_dev.ipynb` exists, copied from `forecast_final.ipynb`, with `WORKBOOK` changed to
  `forecast_dev.xlsx` so it cannot clobber the live file.
- The short-term tool's numbers are **what lands in our bank**: real money, real date, factoring
  already inside them. The notebook does not re-cut them into legs and does not re-date them.
  (Answered by the user while charting.)
- A month under an override is **not** required to add up to budget. That limitation is accepted.
  The Check sheet keeps tying CF to Budget on the budget basis only, as it does today.
  (Answered by the user while charting.)
- The CF basis dropdown has **three** settings: Budget only / Short-term override / Difference.
  (Answered by the user while charting.)
- `Budget input` already carries an `Actuals/Budget` column, at month grain, per country and
  stream. It is set from the notebook constant `FIRST_BUDGET_MONTH = 5` and is **a label only** -
  nothing in the cash model reads it. It marks a P&L amount as closed. It says nothing about when
  that amount reached cash.

## Decisions so far

<!-- one line per resolved ticket -->

- [The short-term input sheet: what it looks like and which weeks are live](./issues/02-input-sheet-shape.md):
  wide wins - `Short-term input` takes a 4x5 rectangle of countries by week, unpivoted beside it
  into the named table `ShortTerm` the CF joins on, with one typed `As of Monday` cell driving the
  window; the risk that moving that Monday alone silently re-keys the figures is accepted, and
  handled by procedure in ticket 04.
- [The pasted week that straddles two months](./issues/08-the-week-that-straddles-two-months.md):
  dissolved by re-keying the whole cash flow at the week level - `cash_week_of()` is now
  `2026-W40`, a week is one column everywhere, and the month is a label over whole weeks. A `CF`
  month therefore no longer matches a `P&L` month, and a `CF` sheet is an ISO year; the `Check`
  sheet ties on the horizon so it is unaffected.

## Not yet specified

- **How the short-term project hands the numbers over.** Ticket 02 fixed what the sheet wants to
  receive: four rows of countries in scope order, five week columns, revenue cash in whole units,
  and no dates inside the block. What is still open is the currency each country's figure is in -
  the CF sheets are DKK throughout and the paste is not obviously converted anywhere - whether the
  tool prints anything beyond revenue, and whether the hand-over is a file, a screen to copy from,
  or a person. <clears-with: 03>
- **Whether anybody needs a true calendar-month cash figure.** Ticket 08 made a `CF` month the
  weeks under that month's label rather than the days in that month, so a `CF` month and a `P&L`
  month no longer agree. Nothing in the workbook depends on their agreeing - the `Check` sheet
  ties on the horizon - but a person might, and collapsing the week groups now gives a monthly
  view that is close rather than exact. Revisit if somebody reads the collapsed view and asks.
- **What the CF charts and `Dashboard CF` do under an override.** They read the same tables, so
  they will move; whether that is wanted is a separate question from the sheets themselves.
- **Scoring the short-term prediction against what happened.** The user hinted at wanting to see
  "how good our prediction is". That is a metrics question, and it needs the actuals basis fixed
  first (ticket 05).
- **Whether the overlay reaches the Budget side at all,** or stays purely a cash-sheet concern.

## Out of scope

<!-- ruled beyond the destination; never graduates -->
