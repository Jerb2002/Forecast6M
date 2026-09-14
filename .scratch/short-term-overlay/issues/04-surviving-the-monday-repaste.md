# Surviving the Monday re-paste

Type: grilling
Status: resolved
Blocked by: 02
Claimed by: claude-opus-5 session 86db36d5
Claimed at: 2026-09-14T20:18:00+02:00
Resolved at: 2026-09-14T21:05:00+02:00

## Question

Section 16 already records the problem: re-running the notebook rewrites the eight P&L sheets and
the three input sheets, so anything typed into them is lost. The weekly overlay makes that fatal
rather than annoying, because the whole point is a paste that happens every Monday and a notebook
that does not run every Monday.

Decide the rule:

1. **Paste, never re-run.** The notebook is a build step; Mondays are Excel-only. Simplest, and it
   means every formula the overlay needs must already be in the written workbook - the notebook
   cannot compute anything on Monday numbers.
2. **The notebook stops overwriting input sheets.** It reads what is there and writes only what is
   missing. Larger change, touches `Factoring`, `COGS delay`, `Cost timing` and `Budget input` as
   well, and turns a known wart into a fixed one.

Option 1 is a constraint on every later ticket, so this is worth settling before anything is built
on top of the input sheet. Whichever wins, say plainly what a CFO does on a Monday, in order.

**Added by ticket 02.** "In order" is now load-bearing rather than tidy. The input sheet is wide, so
a pasted figure carries no date of its own - it knows only which of the five columns it sits in.
Update the `As of Monday` cell and leave the figures alone, and last week's numbers become this
week's silently, with the sheet's own staleness guard reading 0 weeks old. The user accepted that
risk for now rather than have the Monday arrive inside the pasted block. So the procedure this
ticket writes is the only thing standing between a mistyped Monday and a wrong cash flow: say what
happens first, what happens second, and what a CFO should see on the sheet after each step to know
it took.

**Added by ticket 01.** Ticket 06 is closed into this one. The `Actual` label on the P&L sheets
and `Budget input` comes with the CFO's quarterly paste of their budget workbook; the notebook
constant `FIRST_BUDGET_MONTH` is only what a re-run stamps over it. So whichever rule wins here
must say what happens to that label: under option 1 the constant is dead weight the notebook
should stop writing; under option 2 the notebook must leave the pasted column alone. Record which.

## Answer

**Option 1. Mondays are Excel-only, and every re-run of the notebook is a rebuild.**

Two facts settled it before the user was asked anything. The build cell does `book = Workbook()`
and `book.save(WORKBOOK)`; it never opens the file it is about to replace. And nothing in the
notebook computes on amounts: every figure is a share, and the amounts are applied by formulas in
the workbook. So the notebook never needs a Monday number, and option 1 costs nothing on a Monday.

**What a CFO does on a Monday, in order.**

1. Open the live workbook. On `Short-term input`, type the new Monday into `As of Monday`.
   *What shows it took:* the window block reads `this week`, and the five column headers over the
   paste block now show the new week keys.
2. Paste the five weeks of revenue cash over the four country rows. *What shows it took:* the
   figures against the tool, by eye. No new guard was added - option (a) of the four offered; the
   map's Notes say not to add checks nobody asked for, and ticket 02 already accepted this risk as
   a procedural one.
3. Save. The cash flow rolls forward a week. The notebook is not opened.

**Re-runs.** The ticket's option 2 (the notebook stops overwriting input sheets) was declined, and so
was a middle path (the notebook reads the typed cells out of the previous file and seeds them into
the fresh build). A mid-session variant - replace only the four Python Data sheets in the live
file - was ruled out on a fact: every CF formula reaches those sheets by structured reference
(`SUMIFS(PaymentCash[Cash amount], ...)`), and deleting a sheet turns those into `#REF!` for good,
whatever sheet is added afterwards. So the rule is: **a re-run writes a fresh file, and the typed
sheets are carried across by hand.** The user has already tested this for the budget paste. The
checklist is on the README: the eight P&L sheets, `Factoring`, `COGS delay`, `Cost timing`,
`Short-term input`, and the basis switch once ticket 05 places it. Doing the carry-over from the
notebook is not ruled out, only not wanted yet.

**The `Actual` label** (ticket 06, closed into this one). The P&L sheets are not touched; they are
the CFO's paste and carry their own `Actual`/`Budget` row. `Budget input`'s label column was a typed
value with a dropdown, stamped by the notebook off `FIRST_BUDGET_MONTH`, and went stale the moment
a month closed - the paste updates the P&L row and never that column. It is now a **formula off
row 5 of the same P&L the amount beside it reads** (`pl_basis_formula`, the same INDIRECT-and-
ADDRESS as the amount), the dropdown is gone, the column is `Actual/Budget` and the word is
`Actual` on both sheets. `FIRST_BUDGET_MONTH` stays and its comment now says what it is: a
placeholder stamp on a fresh build, like the random figures under it, overwritten by the paste.

Built and checked: `forecast_dev.ipynb` cells for the P&L sheet, the budget sheet and the
settings; the workbook rebuilt and opened in Excel 16 over COM. The column reads `Actual` on 128
rows (January to April) and `Budget` on the rest, carries no validation, and flipping May on
`P&L DNK 2026` to `Actual` re-pointed all sixteen Denmark May rows. Zero error cells anywhere.

**Words for the glossary**, added to `CONTEXT.md`: *rebuild*, *notebook-owned sheet*,
*human-owned sheet*, *Monday paste*.

**What this fixes for ticket 05:** every formula the overlay needs must be in the written
workbook; the notebook cannot help on a Monday. The basis switch cell is human-owned and goes on
the rebuild checklist.

