# Forecast6M

A six-month cash forecast for a four-country business, built in a Jupyter notebook that
writes a single Excel workbook. The notebook takes a monthly P&L budget and turns it into
weekly cash: payment terms, customer and provider mixes, payment delays, work-day rolls,
duty, VAT and factoring legs.

## Files

| File | What it is |
| --- | --- |
| `forecast_final.ipynb` | The live notebook. Writes `forecast_final.xlsx`. |
| `forecast_final.xlsx` | The workbook the CFO reads. |
| `forecast_dev.ipynb` | Development copy. `WORKBOOK` points at `forecast_dev.xlsx`, so it cannot clobber the live file. |
| `full_pre_factoring.ipynb` | Earlier version, kept from before factoring was modelled. |
| `cfo_first_pass_email.md` | The walkthrough that went to the CFO with the file. |
| `requirements.txt` | Python dependencies. |

## Running it

There are two routines, and they never happen on the same day.

### The Monday paste (Excel only)

1. Open the live workbook. On `Short-term input`, type the new Monday into `As of Monday`.
   Check: the window block reads `this week`, and the five headers over the paste block show
   the new week keys.
2. Paste the five weeks of revenue cash over the four country rows. Check the figures against
   the short-term tool by eye. Do it in this order: the figures carry no date of their own, so a
   moved Monday over unpasted figures shifts them a week without saying so.
3. Save. The notebook is not opened.

### A rebuild (the notebook)

Every run of the notebook writes a fresh workbook. Nothing typed into the old file survives, so
after a rebuild carry these across by hand from the previous file, sheet by sheet:

- [ ] The eight P&L sheets, `P&L DNK 2026` through `P&L USA 2027`, including row 5 (Actual /
      Budget). `Budget input` reads both the amounts and the labels off them; nothing to do there.
- [ ] `Factoring`
- [ ] `COGS delay`
- [ ] `Cost timing`, including the month columns
- [ ] `Short-term input`: the `As of Monday` cell and the four rows of five figures
- [ ] `Dashboard CF`: the `Basis` dropdown in C7 (Budget only / Short-term override /
      Difference). A fresh build opens on Budget only.

Then open the file in Excel and check the `Check` sheet still ties. Run the notebook with the
Python that has its dependencies (`py -3.11` on this machine; `requirements.txt` lists them).

## Work in progress: the short-term revenue overlay

Current effort. The goal is a workbook where a CFO pastes five weeks of short-term revenue
cash every Monday, and the cash flow sheets use it for the near weeks while keeping the
budget-driven model for every week after that — without re-running the notebook.

It is planned as a **wayfinder map**: one map file plus one file per open question, under
[`.scratch/short-term-overlay/`](.scratch/short-term-overlay/).

- [`map.md`](.scratch/short-term-overlay/map.md) — the destination, the decisions made so far, and the fog ahead.
- [`issues/`](.scratch/short-term-overlay/issues/) — one ticket per open decision.

### Resuming it

Run this in Claude Code from the repo root:

```
/mattpocock-skills:wayfinder .scratch/short-term-overlay/map.md
```

That loads the map, picks the next unblocked ticket, and works it. Pass a ticket path after
the map to work a specific one instead.
