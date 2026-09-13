# Surviving the Monday re-paste

Type: grilling
Status: open
Blocked by: 02

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
