# The pasted week that straddles two months

Type: grilling
Status: resolved
Blocked by: 02
Claimed by: claude-opus-5 session 5edc48d9
Claimed at: 2026-09-13T17:40:00+02:00

## Question

A cash flow column is not a week. It is a week *of a month* - `cash_week_of()` keys
`2026-09-W37`, and `cash_year_weeks()` deliberately gives a straddling week a column under each
of the two months it touches, so that a month total stays a calendar month and a payroll on the
last working day lands in the right one.

The short-term tool hands over a week. One number, Monday to Sunday, and no month in it.

`Short-term input` currently dates that week by its **Thursday**, which puts the whole amount in
one month's column and leaves the other month's copy of that week empty. That was a decision made
to get the prototype working, not one anybody took, and it is wrong in a way that is invisible:
the week is right, the year is right, the total for the quarter is right, and one month is over
and the next is under by whatever fell on the wrong side of the 1st.

Settle:

- **Does a pasted week get split across the two month-columns, or land whole in one?** Splitting
  needs a rule for the split - work days either side is the obvious one, and it is exactly the
  spread the long-term model already computes in section 3. Landing whole needs only a rule for
  which, and Thursday is a defensible one.
- **If it lands whole, is a month under an override still allowed to be wrong against its own
  budget month?** The map already accepts that a month under an override need not add to budget,
  so this may be inside a limitation that is already taken.
- **Does the answer change what the short-term project is asked to hand over?** A tool that can
  print two part-weeks around a month boundary makes the question disappear.

## Answer

**Neither. The question was dissolved by re-keying the whole cash flow at the week level, and it
was the user who called it** - reading the CF sheets and finding the same week number appearing
twice across the top. A week is now one column everywhere, and the month is a merged label over a
run of whole weeks.

This is wider than the ticket was written for. The ticket asked what to do with a *pasted* week
that straddles a month end; the answer is that no week straddles anything any more, because the
month came off the front of the key.

**What changed.** `cash_week_of()` returns `2026-W40` - ISO week, ISO year, off the Thursday -
instead of `2026-09-W40`. The three other places that built the same string by formula now build
this one: the COGS rows on `Payment date cash` (which already had the answer in the `Week key`
column beside it), the VAT settlement rows, and the delayed branch on `Cost cash`.
`cash_year_weeks()` walks the ISO weeks of a year and groups them under the month each week's
Thursday falls in, so every week has exactly one column and every month still has a group.

**What deliberately did not change.** The facts that spread a cost across a month are still keyed
month-and-week, because a share is inherently a share of a month - a straddling week still needs
one share for September and another for October. That function is now called `month_week_of()`
rather than `cash_week_in()`, so the two ideas stop sharing a word.

**What it costs, stated plainly.** A month total on a `CF` sheet is now the weeks under that
month's label, not the days in that month. A payroll on the 31st can sit in a week labelled with
the month after, and that month's total carries it - so a `CF` month no longer matches the same
month on the `P&L` beside it. A `CF` sheet is also an ISO year rather than a calendar year: the
2026 sheet opens on Monday 29 December 2025.

The `Check` sheet is untouched by all of it. It ties on the whole horizon, summing each sheet's
`Year` column, not month by month - so what matters is that every week is counted exactly once,
and it is.

**Verified in Excel 16 over COM after a full rebuild:**

- 157 week columns over 2026-2028 (52 + 53 + 52), none repeated, and every week a payment or a
  cost lands in has a column.
- `CF DNK 2026`: 53 week keys, 53 distinct, month bands Jan through Dec intact.
- Month totals sum to the year cell exactly - `Cash in total` 100,450,171 either way, difference 0.
- `Check` sheet: all three difference rows nil, verdict `OK`.
- Zero error cells anywhere in the workbook.
- End to end: 1,234 typed on `Short-term input` keys `DNK|Revenue|2026-W37`, and `CF DNK 2026`
  carries a column for `2026-W37`.

Prose in sections describing the old split was rewritten rather than left to rot - the cash flow
section, the cash dashboard section, and the short-term section added by ticket 02.
