# Actuals: does "the past" mean real amounts, or real cash?

Type: grilling
Status: open
Blocked by: -

## Question

The user asked for "actuals for the past" in the cash flow. The workbook already has something
called actuals, and it is not that thing.

`Budget input` carries an `Actuals/Budget` column per country, month and stream. It marks the
**P&L amount** as closed. The cash flow then takes that closed amount and spreads it across weeks
using the model: term mix, customer delay, work-day roll, factoring legs. So a closed month on the
CF sheet today is a real number on invented timing.

Decide which of these "actuals for the past" means:

1. **Real amounts, modelled timing.** What exists today, once the closed boundary is honest. Cheap:
   nothing new to paste, and it follows from ticket 06. But the weekly shape of a closed month is
   still a guess, and the user will be reading those weeks as fact.
2. **Real cash, real dates.** A second paste block: money actually received per country per week.
   Needs a source the user has not named, and raises the question of which rows it covers - revenue
   alone, or cash out too.

The answer fixes the scope of the whole left-hand end of the sheet, and it is what the user's own
question - "this is not month to month, does it need to be?" - is really asking.

Also settle: at what point does a week stop being forecast and become past? A date cell, the
`Actuals/Budget` column, or today's date?

## Comments

### 2026-09-13 - grilling round 1 asked, not answered (session ended)

Claim released. The round below was put to the user and is still open. A later session should
re-ask it rather than re-derive it.

Facts checked in `forecast_dev.ipynb` first:

- Ticket 08 is **implemented**. `cash_week_of()` keys one ISO week, one column. A CF column is a
  week, not a week-of-a-month.
- `FIRST_BUDGET_MONTH = 5` is still a notebook constant, still label-only.
- `cash_rows()` builds `Revenue` / `Revenue duty` / `VAT on sales` as cash-in, a full cost block as
  cash-out, `Net cash flow`, and `Factoring fee` as a memo line below the net.
- `TODAY()` appears 3 times in the workbook already.
- The repo has no `CONTEXT.md`.

**Vocabulary flagged.** The workbook says "actual" for two different things, and this ticket is the
gap between them:

- **Closed month** - a P&L month marked `Actual` in `Budget input`. The *amount* is final. Says
  nothing about when cash arrived.
- **Past week** - a CF week that has already happened.

**Q1 - what does "actuals for the past" mean on a cash flow sheet?**

- **A. Real amount, modelled week.** The month total is real; the notebook still spreads it over
  weeks with terms, delays and factoring. Nothing new to paste. But a past week is still a guess
  being read as fact.
- **B. Real cash, real date.** A second paste block of money that actually hit the bank, per
  country per week. True weeks, but no source has been named, and it covers revenue only - so
  `Cash out` and `Net cash flow` would be part real, part modelled.

Recommended **A**, on the grounds that B needs a source that does not exist, doubles the Monday
paste, and makes `Net cash flow` a mixed number. Under A the sheet must *say* "real amount,
modelled week" so the weeks are not misread.

**Q2 - where does the past stop and the forecast start?**

The CF is weekly now; `Actuals/Budget` is monthly, so a month boundary cannot land inside a week.

- **A. Keep it monthly**, driven by `Actuals/Budget`. Matches the P&L sheets, but moves once a
  month, so for up to 4 weeks the sheet calls weeks "forecast" that are already gone.
- **B. Derive it from `As of Monday`**, the cell ticket 02 already put on `Short-term input`. Past =
  every week before it. No new cell, moves every Monday by itself, and it is at week grain.

Recommended **B**, which also shrinks ticket 06 to a P&L-only job. `TODAY()` was rejected: the sheet
would change under the reader, so a saved or emailed copy would not match.
