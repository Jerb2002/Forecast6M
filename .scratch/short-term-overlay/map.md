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
- ~~The short-term tool's numbers are what lands in our bank, factoring already inside them.~~
  **Overturned in ticket 03.** The tool works from open invoice items and does not know about
  factoring: the paste is expected customer receipts, gross of VAT and duty, at customer timing,
  in DKK. What the sheet does about factoring is ticket 03's answer.
- A month under an override is **not** required to add up to budget. That limitation is accepted.
  The Check sheet keeps tying CF to Budget on the budget basis only, as it does today.
  (Answered by the user while charting.)
- The CF basis dropdown has **three** settings: Budget only / Short-term override / Difference.
  (Answered by the user while charting.)
- `Budget input` already carries an `Actual/Budget` column, at month grain, per country and
  stream. It is **a label only** - nothing in the cash model reads it. It marks a P&L amount as
  closed. It says nothing about when that amount reached cash. ~~Set from the notebook constant
  `FIRST_BUDGET_MONTH = 5`~~ - since ticket 04 it is a formula off the P&L's own row; the constant
  only stamps the placeholder.

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

- [Actuals: does "the past" mean real amounts, or real cash?](./issues/01-actuals-basis.md):
  neither - under *Override* and *Difference* the CF shows nothing for weeks before the first
  pasted week; *Budget only* keeps the whole horizon so `Check` still ties. The boundary is the
  first week header on `Short-term input`, nothing else. The P&L `Actual` label comes with the
  CFO's quarterly paste; ticket 06 closed into 04.
- [Which cash flow rows the override replaces](./issues/03-which-rows-the-override-replaces.md):
  the paste is customer receipts on unfactored invoices, gross of VAT and duty, DKK, revenue
  only. An overridden week splits it by the budget's duty ratio and VAT rate onto `Revenue`,
  `Revenue duty` and `VAT on sales`, then adds the budget's factored legs on top; the memo fee
  keeps its budget figure. The charting premise "factoring already inside" is overturned.
- [Surviving the Monday re-paste](./issues/04-surviving-the-monday-repaste.md): Mondays are
  Excel-only - type the Monday, paste the block, save; the notebook is never run on one. Every
  re-run is a rebuild that writes a fresh file, and the typed sheets are carried across by hand
  off a README checklist. `Budget input`'s `Actual/Budget` column now reads the P&L's own row, so
  the label follows the CFO's paste; `FIRST_BUDGET_MONTH` is a placeholder stamp only.
- [The three-way basis dropdown](./issues/05-the-basis-dropdown.md): one dropdown on
  `Dashboard CF` C7, the workbook name `Basis`, every CF sheet reads it and none has a switch
  of its own. It reaches the sheets through hidden helper rows and a *plain key* the untouched
  rows look up, so only the three replaced rows carry an IF; the IF-on-every-cell version was
  built and timed too and was no slower or faster - both sit on the workbook's 1.7s volatile
  floor. Difference totals stay sums (`Cash out total` is `-`, net is the revenue gap). Past
  and out-of-window weeks hold nought; how they read is ticket 09.

- [What the Check sheet does under an override](./issues/07-the-check-sheet-under-an-override.md):
  pinned to the budget basis - `Check` reads hidden budget-basis rows off the CF sheets, ignores
  `Basis`, and says so in a note cell; no override-gap block (Difference is that); a fourth
  block checks the three VAT rows net to nought over the horizon, which Excel confirms they do
  on the budget basis, and the one verdict tests all three differences. Build is ticket 10.
- [How a past week reads on the cash flow sheet](./issues/09-how-a-past-week-reads.md):
  greyed and blank, header included, one banner in the note cell - and keyed off "no figure
  on this basis" (a new hidden `shown` row), so under Difference the weeks after the window
  grey too. Built and compared as four tabs, then carried into the notebook. A partly-shown
  month totals its shown weeks; no running-balance row; no charts on the CF sheets to break.
  The calc floor is the machine's, not the workbook's: time against `forecast_final.xlsx`.
- [Build the pinned Check sheet and its VAT block](./issues/10-build-the-pinned-check.md):
  built; `Check` OK on four countries under all three bases with the paste in, VAT `Net`
  nought for DNK, NOR and SWE, zero error cells. Eighteen hidden rows, not four: one copy per
  purchase line, because `VAT on purchases` is each line times its own share. Costs two
  seconds on every recalculation, a quarter of the floor; every SUMIFS sits on that floor
  because the tables' key columns chain off `INDIRECT`. The two seconds are accepted for
  now (user, 15 Sep 2026); ticket 11 stays open for whenever it is worth taking.

## Not yet specified

- **Whether anybody needs a true calendar-month cash figure.** Ticket 08 made a `CF` month the
  weeks under that month's label rather than the days in that month, so a `CF` month and a `P&L`
  month no longer agree. Nothing in the workbook depends on their agreeing - the `Check` sheet
  ties on the horizon - but a person might, and collapsing the week groups now gives a monthly
  view that is close rather than exact. Revisit if somebody reads the collapsed view and asks.
- **What the CF charts do under an override.** `Dashboard CF` follows the sheets it reads and
  says which basis it is on (ticket 05). The charts on the `Payment date cash` sheet read the
  tables, not the CF sheets, so they do not move at all; whether they should is a separate
  question from the sheets themselves.

## Out of scope

<!-- ruled beyond the destination; never graduates -->

- [Moving the actual/budget boundary out of the notebook](./issues/06-moving-the-closed-boundary-out-of-the-notebook.md):
  the boundary already arrives with the CFO's quarterly budget paste; what the notebook stamps on
  a re-run is ticket 04's question. Closed into 04.
- **Scoring the short-term prediction against what happened.** Needs bank cash per week, which
  ticket 01 keeps out of the workbook. *Difference* on the dropdown (override vs budget) is what
  this effort gives instead. A fresh effort if a bank-cash source ever appears.
- **Carrying the typed sheets across a rebuild from the notebook.** Ticket 04: by hand for now,
  off the README checklist; the user tested it for the budget paste. Reading the previous file's
  typed ranges into a fresh build is cheap and safe (it is load-and-save that loses charts) but
  not wanted yet. Also ruled out on a fact: replacing only the Python Data sheets in the live
  file, because every CF formula reaches them by table name and a deleted sheet is `#REF!` for good.
- **The overlay reaching the Budget side.** Ticket 01: the P&L sheets are the CFO's paste, the
  CF never shows actuals, so the overlay is a cash-sheet concern only.
