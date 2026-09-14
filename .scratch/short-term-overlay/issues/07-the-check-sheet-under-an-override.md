# What the Check sheet does under an override

Type: grilling
Status: resolved
Claimed by: claude-opus-5 session 5f483705
Resolved at: 2026-09-14T23:30:00+02:00
Blocked by: 05
Assets: .scratch/short-term-overlay/assets/vat_net.ps1

## Question

The user has already set the rule: the checks keep tying CF to Budget on the **budget timeline
only**, exactly as they do now. A month under an override is not expected to add up.

What is not settled is what the Check sheet *shows* when the dropdown is not on Budget only.

- Does it pin itself to the budget basis regardless of the switch, so it always reconciles the one
  thing it can reconcile - and then risks reading as a green tick over a sheet nobody is looking at?
- Or does it gain a second block that states the override gap plainly, so the un-tied amount is
  reported rather than hidden?

The second is close to what the user wants out of *Difference* on the dropdown, so settle whether
these are one feature or two before building either.

Also: the three VAT rows must still sum to nothing over the horizon - what we collect, less what we
pay, is what we settle. An override drops money onto the revenue row without a matching invoice
behind it. Work out whether that identity survives, and if it does not, whether that is a check
which should fire or a check which should exclude overridden weeks.

**Added by ticket 03.** Under override, `VAT on sales` is the VAT carved out of the paste (service
part times the rate) plus the VAT on the budget's factored legs. `VAT settlement` still reads
`VatSettlement`, which is built off `Budget input` and never sees the paste. So the three-row VAT
identity holds only on the budget basis, by construction, not by accident.

**Added by ticket 05.** Built and seen: `Check` reads the visible CF year totals, so it says
*Off - see the difference rows* on all four countries the moment `Dashboard CF` C7 leaves
*Budget only*, and *OK* again when it returns. Under Difference the CF totals *are* the gap,
so `Check`'s own difference rows then read the override gap by accident. If the sheet is to
pin itself to the budget basis, the material is half there: each CF sheet keeps the
budget-basis figure for `Revenue`, `Revenue duty` and `VAT on sales` in hidden rows
(`rev_budget`, `duty_budget`, `vat_budget`, one per week column), but the cost rows have no
such copy - under Difference they are nought. The name `Basis` is what a formula on `Check`
would test.


## Resolution

**Pinned to the budget basis.** `Check` never takes the short-term paste into account, whatever
`Basis` is set to. It reads budget-basis figures off the CF sheets, not the visible totals and
not the Python Data tables: the sheets are what the check proves, and reading the tables would
let a broken sheet formula pass. A note cell under the title says so in one line: *Always on the
budget basis, whatever Basis is set to.*

**No gap block.** *Difference* on the dropdown is the whole answer to "how far off was the
paste"; a second copy on `Check` would restate the same number and drift.

**A VAT block, folded into the one verdict.** A fourth block, *VAT*: collected (`VAT on sales`),
paid (`VAT on purchases`), settled (`VAT settlement`), and a `Net` line that must be nought.
Verified in Excel on Budget only before deciding: the three rows net to exactly 0 for DNK, NOR
and SWE over the whole horizon (USA has no VAT); the later-year CF sheets exist so the last
settlement lands inside it, so no tail exclusion is needed. *Does it tie?* tests three
differences instead of two; the VAT block has no verdict line of its own.

**What the build needs on each CF sheet.** Hidden budget-basis rows keyed on the real cash week,
never the plain key: cost cash over all cost lines, `VAT on purchases`, `VAT settlement`, and
the factoring fee memo. The cash-in copies (`rev_budget`, `duty_budget`, `vat_budget`) already
exist from ticket 05. `Check`'s "Cash flow sheets" rows then read year totals of those hidden
rows in place of the visible totals. Under an override the visible cost rows are nought for past
weeks, which is why a lumped copy will not do for the VAT lines: the check reads them separately.

**Overturned on a fact.** The grilling first assumed VAT could not net to zero over a finite
horizon (last period collected, settled after the end). Excel says it does; see above.

Build is ticket 10.
