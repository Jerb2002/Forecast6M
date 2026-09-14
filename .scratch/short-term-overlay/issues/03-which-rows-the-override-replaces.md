# Which cash flow rows the override replaces

Type: grilling
Status: resolved
Claimed by: claude-opus-5 session 79d896f3
Resolved at: 2026-09-14T08:54:57+02:00
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

## Answer

**The paste is customer receipts on unfactored invoices, gross of VAT and duty, in DKK. In an
overridden week the three cash-in rows are the paste split by the budget's ratios, plus the
budget's own factored legs. The memo fee keeps its budget figure.**

What the paste is, settled with the user and now in `CONTEXT.md`:

- Final revenue per week per country, in DKK, nothing else (no cash out, no cost lines).
- Built from open invoice items, so it is gross of VAT and of duty: it is what the customer pays.
- Factored invoices are **not in it**. The premise recorded while charting ("factoring already
  inside them") was wrong and is struck through on the map. The tool covers open items only; an
  invoice raised after Monday is not in the figure. Weeks 3 to 5 are thinner for it, and that is
  the tool's to fix, not the sheet's.

The rows, for a country and week under *Short-term override*:

1. **Split the paste.** With `r` = budget `Revenue duty` / budget `Revenue` for the CF month the
   week sits under (read off `BudgetInput`), and `v` = the country's VAT rate on the service line
   (read off `VatLines`, nought where there is no VAT):
   `service = paste / (1 + r + v)`, `duty = service x r`, `vat = service x v`. The three add back
   to the paste exactly. Duty carries no VAT, which is why `v` multiplies the service part alone.
2. **Add the factored legs from the budget model.** The paste excludes factored invoices, so the
   budget's `advance` and `retention` legs for that week are added on top of the split: on
   `Revenue`, on `Revenue duty`, and on `VAT on sales` at the share and rate that row applies
   today. Everything except the `direct` leg, in other words. For Norway, Sweden and the US the
   factored share is nought, so this term is nought and nothing is special-cased.
3. **`Revenue`** = service + budget factored legs on the revenue stream.
   **`Revenue duty`** = duty + budget factored legs on the duty stream.
   **`VAT on sales`** = vat + the VAT on those factored legs (the same by-stream lookup the row
   uses today, restricted to the non-direct legs).
   **`Cash in total`** adds up as it does now.
4. **Memo `Factoring fee`** keeps its budget figure in an overridden week. The fee belongs to the
   same factored legs that were just added, so the memo still describes money that is on the sheet.

The courier / freight-forward split does not arise: the cash model already adds the two into one
`Revenue` stream before anything reaches a CF sheet, so the sheet has nothing to split them into.

Rejected on the way:

- Reducing the paste for factoring (fee only, or fee plus advance). Moot once it was clear that
  factored invoices are not in the paste at all.
- Putting the whole paste on `Revenue` and blanking the other two rows. Total right, rows wrong,
  *Difference* useless row by row.
- A separate `Short-term receipts` row. Three rows that mean the same thing under both bases are
  worth more than one row that means something new.
- Asking the short-term project to apply factoring themselves. Not needed: the budget model has
  the legs and the paste cleanly excludes them.

Ticket 05 builds this; the formulas are two lookups per cell where there is one today (the paste
key plus a by-leg `SUMIFS` on the payment table), which is what its recalculation test is for.
Ticket 07 inherits the VAT identity question with `VAT on sales` now partly paste-derived.
