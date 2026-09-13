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
