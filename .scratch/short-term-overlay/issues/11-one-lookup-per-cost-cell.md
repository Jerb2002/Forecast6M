# One lookup per cost cell, or two

Type: grilling
Status: open
Blocked by: 10

**Parked.** The two seconds are accepted for now (user, 15 Sep 2026). Open, not urgent.

## Question

Ticket 10 gave every purchase line a hidden budget-basis copy keyed on the real week, so
`Check` can tie the cost side and the VAT whatever `Basis` says. Each of those fourteen
copies is the same SUMIFS the visible row already carries on the plain key, so the cost
rows are now looked up twice per week column: about nine thousand extra lookups across the
twelve sheets, and two seconds on every recalculation on today's machine (a quarter of the
floor, which every SUMIFS in the workbook sits on because the tables' key columns chain off
`INDIRECT`).

Settle whether the visible cost rows should read their hidden copy times `mult` instead,
which is one lookup per cell and a multiply, and leaves the plain key to the two VAT rows
and the settlement.

- Ticket 05 kept the untouched rows untouched because the clock showed no difference between
  the two ways of reaching them. This clock shows one. Does that overturn the choice, or is
  two seconds on a nine-second floor nobody's problem?
- If the cost rows fold, the plain key is used by three rows out of thirty; is it still
  worth the row, or do those three take an IF like the cash-in rows and the plain key goes?
- Either way: measure. Build it as the B variant on a copy and clock both with
  `time_basis.ps1`, the way ticket 05 did. Do not settle it by argument.
