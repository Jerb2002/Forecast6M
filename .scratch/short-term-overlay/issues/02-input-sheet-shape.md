# The short-term input sheet: what it looks like and which weeks are live

Type: prototype
Status: resolved
Blocked by: -
Claimed by: claude-opus-5 session 5edc48d9
Claimed at: 2026-09-13T16:26:41+02:00
Assets: .scratch/short-term-overlay/assets/st_code_both_variants.py

## Question

Where does a CFO paste five weeks of revenue cash per country, and how does the sheet know which
five weeks those are?

Two things to settle together, because the shape decides how the window can work:

**Shape.** Wide - a row per country, five week columns across - matches how a human pastes and
reads. Long - country, week key, amount, one row each - matches every other input sheet in this
workbook and is what `SUMIFS` on a composite key already wants. The CF sheets join on
`country|line|week`, so a long table needs no reshape and a wide one needs a helper block.

**Window.** Either one "as of Monday" date cell drives the five week keys and the paste columns
are unlabelled positions, or the week keys are typed and the sheet checks them. The first is one
cell to update and silently wrong if nobody updates it. The second is five cells to update and
cannot drift without showing it.

**Staleness.** What the sheet does when the paste is from three Mondays ago. Silently applying it
is the failure mode that matters here, because the numbers will look plausible.

Build both shapes as lettered variants in `forecast_dev.ipynb`, write the workbook, open it in
Excel and compare. Do not settle this by argument.

## Answer

**Wide, on a sheet called `Short-term input`, with one typed Monday driving the window.**

Both shapes were built into `forecast_dev.ipynb` as `Short-term A` and `Short-term B`, written to
`forecast_dev.xlsx`, and driven in Excel 16 over COM: numbers typed in, keys read back, the anchor
rolled forward, the join checked. The user picked A on one ground, and it is the ground that
decides this sort of question: *end users do not think in long tables*. A twenty-row table is the
shape the rest of the workbook is built in and the shape the cash flow wants, and it is still the
wrong thing to put in front of the person who has to paste into it every Monday.

**The shape.** Four rows of countries, five week columns across. The paste block is a styled range
rather than an Excel table, because a table will not hold a formula in its header and these headers
have to show the live week. Beside it, out of the way, the same figures unpivoted into twenty rows
carrying `Country`, `Week key`, `Cash key` and `Amount` - the named table `ShortTerm`, in the one
shape a CF `SUMIFS` can read. Every cell of it is a formula back into the rectangle, so the two
cannot disagree. That block is what the wide shape costs, and it is the whole of what it costs.

**The window.** One entry cell, `As of Monday`. The five Mondays, the five week keys and the five
column headers are all derived from it, so the window cannot be half-updated. The week key is
`YYYY-MM-Www`, built off the **Thursday** of each week to match `cash_week_of()`, and counted with
`INT((thursday-DATE(YEAR(thursday),1,1))/7)+1` rather than `ISOWEEKNUM` - openpyxl writes an Excel
2013 function without the `_xlfn.` prefix it needs, and the first build opened on `#NAME?` in every
key cell. The format strings are digits only for the same class of reason: `yyyy` and `mm` inside
`TEXT` are spelt differently in a Danish Excel.

**Staleness, and the risk this shape accepts.** The prototype separated two failures that had been
one:

1. *Nobody touched the sheet.* The Monday is old and so are the figures. Caught: the window block
   reads `3 weeks old - repaste`.
2. *The Monday was updated and the figures were not.* In a wide block a figure knows only which
   column it is in, so last week's numbers silently become this week's - and `Weeks old` now reads
   0, so the one guard reports everything as fine. Driven in Excel: a 1,000 keyed
   `DNK|Revenue|2026-09-W37` re-keyed itself to `2026-09-W38` on the anchor moving, with no
   complaint. The long shape caught this per row (`OUT OF WINDOW`); the wide shape cannot.

Failure 2 is the one the ticket called "the failure mode that matters", and the chosen shape cannot
see it. The fix that would close it is to have the Monday arrive *inside* the pasted block, so date
and figures always move together; the user was asked and chose **a separately typed date for now**,
accepting the risk. What is left is procedural: the note on the sheet says to set the Monday first
and paste second, in that order, and says why. Ticket 04 owns the Monday procedure and now carries
this as a constraint.

**What was built.** One markdown cell and one code cell in `forecast_dev.ipynb`, just before the
build cell, and `write_short_term(book)` added to the Inputs group. Variant B was deleted from the
notebook once the choice was made; its code is parked in the asset above. The workbook opens in
Excel 16 with zero error cells anywhere, and `Short-term input` is tab 32 of 37, under
`Inputs -->`, and listed on Contents.

Nothing reads the sheet yet. A small proof block on the right runs the real CF `SUMIFS` against
`ShortTerm` for week 1, so the join is visible rather than assumed; ticket 05 is what makes it real.
