# Six-month cash forecast — first pass

**Subject:** Cash forecast model — first pass, and what I'd like to agree in our meeting

---

Hi,

Attached is the first working version of the cash forecast. The structure is finished and
every calculation runs end to end — what is not finished is the **numbers**, which are
placeholders throughout. I have listed at the bottom exactly which ones, so nothing in here
gets mistaken for a figure we stand behind.

What I would like out of the meeting is agreement on the variables. Once those are set, the
model produces the weekly cash flow on its own.

**The one thing worth knowing before you open it:** you budget the way you already budget —
month by month, per country, on a normal P&L. Everything else follows automatically from
that. You never enter a cash figure anywhere.

---

## How to find your way around

Open the file and you land on **Contents** — every sheet in the workbook, grouped and
linked. Each sheet links back to it from its top-left corner. Nothing is more than one
click away.

The groups run left to right in the order the numbers move through them. Coloured tabs mark
where one group ends and the next begins.

**Blue means you can type in it. Black means the model worked it out.** That holds on every
sheet in the file.

---

## What each section contains

### 1 & 2 — Budget 2026 and Budget 2027

Eight sheets: one per country, per year. **This is your budget, in your own format** —
months across, lines down, in DKK thousands.

Revenue and COGS split into Courier, Freight Forward and Duty; then Manpower, then SG&A line
by line — Staff cost, Vehicle, Rent, IT, Sales & Marketing, Travel & Rep, Admin & Misc, Bad
debt, SLA, External Assistance; then the usual subtotals down to EBITDA, EBIT and net
profit, plus the gross margin and EBITDA margin ratios.

2026 is marked Actual for January to April and Budget from May; 2027 is Budget throughout.
Subtotals and margins are formulas — you type the lines, not the totals.

**These eight sheets are the only place the business's own figures go.** Everything else in
the workbook reads from them.

### 3 — Cash flow

Twelve sheets: one per country for 2026, 2027 and 2028, **week by week under its month**.

Cash in — revenue, revenue duty, VAT on sales. Cash out — COGS, COGS duty, each manpower
and SG&A line separately, VAT on purchases, and the VAT settlement. Then net cash flow for
the week. The factoring fee sits below the net as a memo, inside no total: the bank keeps
it out of what it pays us, so the revenue line is already net of it and nothing is ever
paid out to settle it.

In front of those twelve is **`Check`**, which is where I would start. It puts each
country's budget beside its own cash flow sheets, pre-VAT, and bridges the two: the
factoring fee the bank keeps, the lines that never reach cash, and the years the tail of
the horizon lands in. What is left is the difference, and it is nought in all four
countries.

**Nothing is typed on these sheets.** Every cell is a lookup. They are the output.

2028 is there because an invoice raised late in 2027 on 60-day terms, paid late, lands in
early 2028 — and a forecast that dropped it would understate nothing but would simply lose
the money.

### 4 — Inputs

The sheets that decide **when** money moves, as opposed to how much.

| Sheet | What it holds | Typed? |
|---|---|---|
| **Budget input** | Your P&L in long form — one row per country, month and line. A bridge between the P&L sheets and the calculations. | No, formulas |
| **Factoring** | Which countries sell invoices to the bank, in which months, and on what share, advance rate and fee | **Yes** |
| **COGS delay** | How many days late we actually pay each courier, by the month its invoice falls due. Buttons beside the table filter it to the row you want | **Yes** |
| **Cost timing** | One row per country and cost line: the rule it normally pays on, then a column per month for the days it actually ran late | **Yes** |
| **Dashboard** | Pick a country and a year and read the P&L, with a second country beside it to compare | No, formulas |

**Cost timing** is the sheet worth two minutes of the meeting. Each of the twelve cost lines
pays on one of four rules — last working day of the month, the 15th, spread evenly across
the month's working days, or no cash effect at all (which is what Bad debt is). The months
across the sheet are where you say *"the rent waited three weeks in March"*, and the cash
flow moves accordingly — April then carries March's rent as well as its own.

### 5 — Python Data

Six sheets the script writes and rewrites on every run: the working-day calendar, the
invoices by due date, the dated costs, the dated payments, the VAT rules and the VAT
calculation.

**Nothing here is for editing** — it is the working that the cash flow sheets read. It is
visible because a forecast whose workings are hidden is a forecast nobody can check.

---

## How a budget figure becomes cash

This is the whole model in seven steps, and all of it is live Excel formulas:

1. You type a figure on a P&L sheet.
2. **Budget input** picks it up automatically, for that country, month and line.
3. Revenue and COGS are spread across the month's working days, given a payment term from
   the customer and provider mix, and turned into a **due date**.
4. The **delay** for that courier and month is added, and the result rolled forward onto a
   working day against that country's own bank holidays. That is the **payment date**.
5. Manpower and SG&A take their date from their rule on **Cost timing**, plus whatever delay
   that month carries.
6. **VAT** is charged on the invoices in and out, and settled on each country's own filing
   clock — monthly in Denmark and Sweden, bi-monthly in Norway, none in the US.
7. Where factoring applies, the invoice is split into the bank's advance and the retention
   that follows when the customer eventually pays.

The **cash flow sheets** then add all of it up, week by week.

Because every step is a formula, changing a budget figure — or a delay, or a timing rule —
moves the cash flow **immediately**. The script does not need to run again for any of it.

---

## What you can change, and what the script decides

### Change directly in Excel

- **The budget** — all eight P&L sheets
- **Factoring** — countries, months, factored share, advance rate, fee
- **COGS delay** — days late per courier, country and due month
- **Cost timing** — the rule per cost line and country, and the days late per month

### Decided by the script (shown in the workbook, not editable there)

- **Scope** — Denmark, Norway, Sweden, United States; 24 months from January 2026
- **The working-day calendar and bank holidays** for each country
- **How a month's invoicing is spread** across its working days — currently even
- **The payment terms on offer** (Cash through Net +60, and current-month variants) and the
  **customer mix** across them
- **The providers** we buy COGS from and the split between them — currently 80/20
- **Which P&L row each cost line is budgeted on**
- **VAT** — rates, filing frequency, which lines carry it and how much is reclaimable
- **Factoring mechanics** — the weekly invoice run day, when the bank settles it, and how
  long the retention trails the customer's payment

These are all one-line changes on our side. They sit in the script rather than the workbook
because they are rules rather than figures anybody revises month to month — and because a
rule typed in four places eventually disagrees with itself.

---

## Please treat every number in it as a placeholder

To be completely clear about what is and is not real in this version:

- **The P&L figures are generated, not real.** They exist so the mechanics can be seen
  working. They are the first thing your budget replaces.
- **The customer term mix is a placeholder** — it needs the real split.
- **Cost delays start at zero everywhere**, i.e. everything pays on its rule.
- **The VAT filing frequency was inferred from the placeholder revenue**, so Denmark's in
  particular needs confirming — quarterly rather than monthly moves January's VAT payment
  from February to June.
- **Bank holidays are rule-generated** and want checking against a real source per country.
- **Every P&L is in DKK thousands**, whatever the country. The FX rate is applied before
  anything is typed, and the workbook does not currently record which rate was used.

---

## What I would like to agree in the meeting

1. **The budget itself** — 2026 and 2027, per country
2. **Customer payment terms** — the real mix per country
3. **Providers** — who they are, their terms, and the split between them
4. **Actual payment delays** — what we run at on both sides, and where it varies by month
5. **VAT** — the filing frequency per country, whether the Vehicle line is vans or company
   cars, and whether any of the rents are on voluntarily registered landlords
6. **Factoring** — which countries, which months, and the commercial terms
7. **Cost timing** — the rule for each of the twelve cost lines, per country
8. **FX** — the rate each country's budget was translated at, and where that should live
9. **The invoicing profile** — is even across working days right, or is there a month-end
   skew we should model?

Happy to walk through it live and change things as we go — most of the above is a single
cell.

Best,
Justo
