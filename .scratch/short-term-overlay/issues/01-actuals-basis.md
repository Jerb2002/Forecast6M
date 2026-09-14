# Actuals: does "the past" mean real amounts, or real cash?

Type: grilling
Status: resolved
Blocked by: -
Claimed by: claude-opus-5 session 0964d4b6
Claimed at: 2026-09-14T08:40:00+02:00

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

### 2026-09-14 - resolved

**Neither A nor B. Past weeks are not shown in the cash flow at all.**

The user's words: "we will not be showing in the cash flow model the actuals at all ... for
anything that comes before [the five weeks], we will have that as blank or a message ... so that
any user knows the forecast does not apply to those periods."

What that fixes:

- **No actuals basis in the CF.** There is no second paste of bank cash, and no re-labelled
  "real amount, modelled week". The left-hand end of the sheet under an override is empty or
  marked *past*, and nothing else.
- **Only under Override and Difference.** *Budget only* keeps every week of the horizon as it
  does today, so the `Check` sheet still ties CF to Budget on the budget basis. (Q3, accepted.)
- **The boundary is the first week of the paste block on `Short-term input`.** Past = every CF
  week whose key sorts before that first header. Live = that week and everything after. No other
  cell, and not `TODAY()`. The header is worked out from `As of Monday` as ticket 02 built it;
  the user's "no" to naming the cell was about which thing is the source of truth (the block's
  first week), not about how it gets there. Ticket 02 stands.
- **The `Actual` label on the P&L side is the CFO's.** The eight P&L sheets and `Budget input`
  are paste targets mirroring the CFO's own budget workbook, refreshed roughly quarterly, and
  the actual months come in with that paste. The notebook constant `FIRST_BUDGET_MONTH` only
  matters for what the notebook *writes over* on a re-run - which is ticket 04's question. Ticket
  06 is closed into 04.

What it opens:

- How a past week *reads* on a CF sheet (blank cells, a marker, greyed columns) and what a month
  subtotal does when its weeks are partly past: ticket 09, prototype, after 05.
- Scoring the prediction against what happened needs bank actuals, which this decision keeps out
  of the workbook. Moved to Out of scope on the map.
