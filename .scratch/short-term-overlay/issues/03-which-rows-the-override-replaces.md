# Which cash flow rows the override replaces

Type: grilling
Status: open
Blocked by: -

## Question

The short-term tool gives one figure: revenue cash per country per week. The cash flow sheet has
three rows that move with revenue - `Revenue`, `Revenue duty`, and `VAT on sales` - plus a memo
`Factoring fee` line below the net.

The override number is what lands in the bank, so factoring is already inside it. Decide what
happens to the other three rows in an overridden week:

- Does the one figure cover revenue **and** its duty, or revenue alone with duty left modelled?
- Is it gross of VAT, so `VAT on sales` must be carved out of it, or net, so that row stays
  modelled and may now double count?
- The memo fee line reads the payment sheet by leg. An overridden week has no legs. Does the memo
  go blank for those weeks, stay modelled, or drop out of the basis entirely?

`VAT on sales` is the trap. It is the one row on the sheet that reads the payment table by stream
rather than by the composite key, precisely because a factored invoice settles on four legs and
only three are revenue. An override that arrives already net of the bank fee breaks the assumption
that row was built on.

Get this wrong and the sheet is wrong by roughly the VAT rate in every overridden week - which is
large, and looks like nothing.
