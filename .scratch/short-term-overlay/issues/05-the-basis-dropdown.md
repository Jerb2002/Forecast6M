# The three-way basis dropdown

Type: prototype
Status: open
Blocked by: 02, 03

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
