# Build the pinned Check sheet and its VAT block

Type: task
Status: resolved
Claimed by: claude-opus-5 session d6eda8f5
Resolved at: 2026-09-15T00:05:00+02:00
Blocked by: 07
Assets: .scratch/short-term-overlay/assets/patch_check.py, drive_check.ps1, 10-built-Check-Short-term-override.png

## Question

Carry ticket 07's decision into `forecast_dev.ipynb`, one section at a time:

1. On every CF sheet, hidden budget-basis rows keyed on the real cash week: cost cash over all
   cost lines, `VAT on purchases`, `VAT settlement`, factoring fee. Beside the existing
   `rev_budget`, `duty_budget`, `vat_budget`.
2. `Check`'s "Cash flow sheets" rows read year totals of those rows instead of the visible totals.
3. A fourth block, *VAT*: collected, paid, settled, `Net`. The verdict tests three differences.
4. A note cell under the title: *Always on the budget basis, whatever Basis is set to.*

Done when, in Excel via COM, `Check` reads OK on all four countries under all three settings of
`Basis`, with the paste in, and the VAT `Net` line is 0 for DNK, NOR and SWE. Rebuild the workbook
and time a full calc against ticket 05's 1.7s floor.

**Added by ticket 09.** `CF_BASIS_ROWS` now has fourteen rows (`shown` after `mult`), and
the greying in `grey_unshown_weeks()` ranges from `CF_FIRST_ROW` to `max(CASH_ROW_AT)`, so
hidden rows added to `CF_BASIS_ROWS` shift everything and stay outside the greyed block by
themselves. The 1.7s floor is not on this machine today: `forecast_final.xlsx` takes 7s per
Calculate. Time the build against that file on the day, with ticket 05's `time_basis.ps1`.


## Resolution

**Built, and the done condition holds.** `Check` reads *OK* on all four countries under all
three settings of `Basis`, with the paste in; the VAT `Net` line is 0 for DNK, NOR and SWE
(USA has no VAT and reads `-`); zero error cells across the workbook; the note cell under
the title reads *Always on the budget basis, whatever Basis is set to.*

**What went into the notebook** (`patch_check.py`, one pass over the CF cell and the Check
cell, plus their write-ups):

1. *The hidden rows.* Not four but eighteen. `VAT on purchases` is each cost line times its
   own VAT share, so a lumped cost figure cannot give it back; each of the fourteen purchase
   lines (COGS, its duty, twelve scheduled lines) gets a copy keyed on the real week, with
   the line's VAT share beside it in the hidden column, and the four summary rows sit under
   them: `cost_budget` (SUM of the block), `vat_out_budget` (the same SUMPRODUCT the visible
   row is, over the block), `settle_budget`, `fee_budget`. The purchase list is read off
   `CASH_ROWS`, so a cost line added in section 11 gets its copy by itself. `CF_BASIS_ROWS`
   moved below `cash_rows()` for that; `CF_FIRST_ROW` is now 33 and the visible `Revenue`
   row 42. The seven rows `Check` reads carry month and year totals like the visible rows.
2. *`Check` reads them.* `cf_year_terms()` now indexes `CF_BASIS_ROW_AT`: cash in is
   `rev_budget + duty_budget`, the fee is `fee_budget`, cash out is `cost_budget`.
3. *The VAT block.* Collected, paid, settled, `Net` (rounded like the other differences).
   The verdict tests three differences; the status formula takes any count.
4. *The note.* `write_note(sheet, CHECK_NOTE)` under the title.

**Checked in Excel 16** (`drive_check.ps1`): on Budget only, each hidden copy equals its
visible row in the year column of `CF DNK 2026` (`cost_budget` 67,350,000 plus VAT on
purchases 11,530,091 plus settlement 3,477,250 is the visible `Cash out total` 82,357,341,
exactly), and the copies hold still under Override and Difference while the visible rows
move. Picture of the sheet under Override in the assets.

**Timing, back to back on the same Excel, same afternoon** (`time_basis.ps1`):

| step | before ticket 10 | after | `forecast_final.xlsx` |
|---|---|---|---|
| full recalculation | 13.1s | 15.8s | 11.4s |
| Calculate, no edit (the floor) | 8.1s | 9.9s | 6.7s |
| switch basis, any direction | 9.0s | 10.9s | - |
| edit one pasted figure | 7.4s | 9.5s | - |

About two seconds on every step, a quarter. The build itself took 17 minutes in openpyxl.

**Why lookups sit on the floor at all.** Every Python Data table's key and amount columns
are formulas chained off `Budget input`, which reads the P&Ls through `INDIRECT`. So every
SUMIFS in the workbook is a dependent of a volatile cell and recomputes on every Calculate,
and the floor scales with the count of lookups, not with what changed. That is why ticket
05's six copies were free on a 1.7s floor and ticket 10's eighteen are not on a 8s one.
Pre-existing; not this ticket's to fix.

**Handed on as ticket 11.** The extra cost is the purchase block doing the same fourteen
lookups the visible cost rows already do. The visible cost rows could read their hidden copy
times `mult` instead, one lookup per cell rather than two, and the plain key would then be
needed on the two VAT rows only. Ticket 05 kept the untouched rows untouched on a clock that
showed no difference; this clock shows one, which is a new fact and a decision to make, not
one to slip in under a build ticket.
