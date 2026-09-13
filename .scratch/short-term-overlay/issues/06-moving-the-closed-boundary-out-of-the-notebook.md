# Moving the actual/budget boundary out of the notebook

Type: task
Status: open
Blocked by: 01

## Question

`FIRST_BUDGET_MONTH = 5` is a notebook constant. It sets which months are labelled `Actual` on the
eight P&L sheets and on `Budget input`. Every month that boundary moves forward by one, and today
that means editing the notebook and re-running it - which, under ticket 04, may be the one thing
the Monday workflow must not require.

This is the user own question from charting: *"at the moment this is not month to month, does it
need to be?"*

Do the work to make the boundary an input rather than a constant, and record what it turned out to
cost. Which input it becomes depends on ticket 01: a single "closed through" date cell if actuals
stay a P&L label, or the pasted `Actuals/Budget` column itself if the user is going to maintain it
row by row when they paste a new budget.

Nothing downstream reads the column today - it is a caption on the dashboard and a label on the
P&L. If ticket 01 makes the cash model read it, say so here, because that turns a label into a
load-bearing input.
