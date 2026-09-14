# How a past week reads on the cash flow sheet

Type: prototype
Status: resolved
Claimed by: claude-opus-5 session 568f71e6
Resolved at: 2026-09-14T23:40:00+02:00
Blocked by: 05
Assets: .scratch/short-term-overlay/assets/variants_09.ps1, snap_09.ps1, 09-variants.xlsx (four variant tabs, not committed), 09-09-*.png (the variants), patch_shown.py, drive_shown.ps1, 09-built-*.png (the build)

## Question

Ticket 01 decided that under *Short-term override* and *Difference* the CF shows nothing for
weeks before the first pasted week: "blank, or a message ... so that any user knows the forecast
does not apply to those periods". Under *Budget only* those weeks stay as today.

Build the options in the real workbook and compare them. Do not settle by argument.

- **A. Empty cells.** The columns are there, the figures are gone. Cheapest, and the least
  informative: an empty column and a zero-cash week look the same at a glance.
- **B. A marker in every cell.** `past`, `-`, or similar, in the body cells. Loud; may fight with
  the number formats and the row totals.
- **C. Greyed columns with one banner.** Body cells empty, the column header or a row above it
  says *past* once, the fill is grey. One thing to read, no per-cell text.

Settle at the same time:

- What a **month subtotal** shows when its weeks are partly past: the sum of the live weeks, or
  blank/marked like the weeks under it.
- What **`Net cash flow`** and the running balance do across the boundary: start from zero at the
  first live week, or from an opening balance typed somewhere. (If the second, that is a new cell
  and belongs to ticket 04's Monday procedure.)
- Whether the **charts** on the CF sheets survive an empty left end without redrawing wrongly.

Link the variant workbooks as assets. Open them in Excel via COM, not openpyxl.

**Added by ticket 05.** What is there to build on:

- A past week holds **0 on every row**, not an empty cell, so the `+`/`-` formulas in the
  totals and the net keep working. Under `CF_MONEY` that reads `-`. Blank-as-in-empty would
  need the totals rewritten as `SUM`s first; blank-as-in-shown (a `;;;` format keyed off a
  flag) would not.
- The hidden row `Live on this basis` (row 8, one 1/0 per week column) is the per-column flag
  a conditional format, a grey fill or a banner can key off. `Overridden by the paste` (row 9)
  marks the five pasted weeks the same way.
- A month whose weeks are partly past currently **totals its live weeks**. Nothing decides
  that; it is what `SUM` does.
- `Net cash flow` is nought in past weeks. There is no running balance on the sheet today, so
  the "opening balance" question in this ticket is about adding one, not about where it starts.
- Under *Difference*, the weeks **after** the five pasted ones also come to nought on every
  row and read `-`, the same as a past week and the same as a week with no cash. The user
  asked for that look to be decided here, together with the past weeks, rather than in 05.


## Answer

**Greyed, blank, header included, with one banner in the note cell - and the rule is "no
figure on this basis", not "past".** Built as four tabs on a copy of the workbook, compared
in Excel, then carried into `forecast_dev.ipynb` and checked there.

**What was compared.** Four copies of `CF DNK 2026` with the paste in, each a conditional
format keyed off the hidden helper rows, nothing in the figures changed:

- A, empty cells: `;;;` where `live` is 0. Cheapest; an unshown column and a no-cash week
  both read as nothing, and the eye cannot tell them apart.
- B, a marker: `past` in every body cell, totals included. Loud, and the word is wrong under
  Difference, where the weeks *after* the window are just as empty.
- C, grey and a banner: body `;;;` plus the `GRID` fill, the week header greyed with `MUTED`
  text, and the note cell saying *weeks before 2026-W38 are not shown*. Keyed off `live`,
  so under Difference the weeks after the five read `-` like a no-cash week.
- D, C's look keyed off "no figure on this basis" (`mult` = 0 and `override` = 0): under
  Difference only the five pasted weeks show, and the banner says so.

The user chose D. C and D are the same thing under Override; the choice between them was
the Difference question ticket 05 handed on, and it went the same way as the past weeks: a
week with nothing to say is greyed, whichever side of the window it is on.

**The three side questions.**

- *Month subtotal, partly past.* Sum of the shown weeks, which is what `SUM` does and what
  was there already. A month with no shown week is greyed with its weeks. Sep 2026 in the
  picture: W36-37 grey, W38-39 live, Total = W38 + W39.
- *Net cash flow and a running balance.* A `Cumulative net` row (from nought at the first
  shown week) was built on all four tabs and the user dropped it: no running balance on the
  sheet, so no opening-balance cell and nothing new for ticket 04's Monday checklist.
- *Charts.* There are none on the CF sheets; the CF charts sit on `Payment date cash` and
  read the tables. Nothing to survive. (Still fog on the map: whether those should move.)

**What is in the notebook now.**

- A fourteenth hidden row, `shown` (*Figure on this basis*): 1 when `mult` or `override`
  is 1. The greying keys off it; a banner or a dashboard can too.
- Per CF sheet, per month: one rule over the month's run of week columns
  (`{first}$11=0`, relative) and one over its total column (`SUM` of the run's `shown` cells
  = 0), each in two ranges - the header row with grey fill and `MUTED` text, the body rows
  with grey fill and the `;;;` number format. The year column is never greyed. 48 rules a
  sheet, 576 in the workbook.
- The note cell: `Basis: Short-term override   |   weeks before 2026-W38 are not shown`, or
  `Basis: Difference   |   only the pasted weeks are shown`, or `Basis: Budget only`.
- `CONTEXT.md`: *unshown week*.

**Checked in Excel 16** (paste in, all twelve CF sheets): zero error cells; `Check` OK on
all four countries under Budget only; the week before the paste greys on the four 2026
sheets under Override and on all twelve under Difference, and on none under Budget only;
the pasted weeks and the year column never grey; the fill reads back as 14278881 (`GRID`).

**Timing, and a fact for ticket 10.** The rules cost nothing measurable: the same file with
all 576 stripped clocks the same (13.2s full / 8.1s floor with, 17.6s / 11.8s without, run
to run). But the floor itself has moved: `forecast_final.xlsx`, which this effort never
touches, takes 12s full / 7s per Calculate in the same Excel today, against ticket 05's
2.7s / 1.7s. The machine, not the workbook. Ticket 10's "time against the 1.7s floor" should
time against `forecast_final.xlsx` on the day, not against the number.
