"""The two markdown paragraphs that describe the split week, rewritten for the new keying."""
import io, json

NB = "forecast_dev.ipynb"
nb = json.load(io.open(NB, encoding="utf-8"))
cells = nb["cells"]


def patch(index, old, new):
    src = "".join(cells[index]["source"])
    assert src.count(old) == 1, f"cell {index}: {src.count(old)} matches"
    src = src.replace(old, new)
    parts = src.rstrip("\n").split("\n")
    cells[index]["source"] = [p + "\n" for p in parts[:-1]] + [parts[-1]]


patch(63, """Each month's weeks sit under it and that month's own total sits at the end of them. The
week columns are grouped, so the buttons above them collapse a month down to its total and
the sheet becomes a monthly view without needing a second tab. A week straddling two months
is a column under each, holding only that month's days, so a month total is a real calendar
month and ties to the budget month behind it.""",
"""Each month's weeks sit under it and that month's own total sits at the end of them. The
week columns are grouped, so the buttons above them collapse a month down to its total and
the sheet becomes a monthly view without needing a second tab.

**A week is one column.** It was two for a while - the month on the front of the key, so a
week straddling a month end appeared under each month holding only that month's days. That
bought a month total that was a real calendar month and tied to the budget month behind it,
and it cost the thing the sheet is read for: the same week number twice across the top of a
sheet whose columns are weeks. The month is now a label over a run of whole weeks. A week
belongs to the month its Thursday falls in, which is the same Thursday that already decides
the week number and the year.

What that gives up is worth knowing. A month total here is the weeks under that month's
label, not the days in that month, so a payroll on the 31st can sit in the month after it
and a `CF` month no longer matches the same month on the `P&L` beside it. The `Check` sheet
is unaffected, because it ties on the whole horizon rather than month by month, and every
week is still counted exactly once. A `CF` sheet is also an ISO year rather than a calendar
one: 2026 opens on Monday 29 December 2025.""")

patch(65, """**The one thing that is genuinely harder than on the budget dashboard** is which column to
read. A P&L's January is always column C. A cash flow sheet's January is a run of week
columns and then a total, and how many weeks a month has depends on the calendar - a week
straddling two months is a column under each. June's total is in `AO` on the 2026 sheets and
in `AN` on the 2027 ones.""",
"""**The one thing that is genuinely harder than on the budget dashboard** is which column to
read. A P&L's January is always column C. A cash flow sheet's January is a run of week
columns and then a total, and how many weeks a month is labelled with depends on the
calendar - four in most months, five in some - so a month's total column moves from year to
year.""")

json.dump(nb, io.open(NB, "w", encoding="utf-8", newline="\n"), indent=1, ensure_ascii=False)
print("prose updated in cells 63 and 65")
