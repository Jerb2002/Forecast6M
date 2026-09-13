"""Re-key the cash flow at the week level: one column per ISO week, month as a label only.

Four sheets build a `Cash week` string and every one of them put the calendar month on the
front, so a week straddling a month end became a column under each month. That kept a month
total a real calendar month and cost the thing the sheet is actually read for: the same week
number appearing twice across the top.

After this, `Cash week` is the plain ISO week - `2026-W40` - everywhere, and a month is a
merged label over a run of whole weeks. The facts that spread a cost across a month keep
their month, because a share is inherently a share of a month; they are renamed so the two
ideas stop sharing a word.
"""
import io, json

NB = "forecast_dev.ipynb"
nb = json.load(io.open(NB, encoding="utf-8"))
cells = nb["cells"]


def patch(index, old, new, count=1):
    """Replace exactly `count` occurrences in one cell, or fail loudly."""
    src = "".join(cells[index]["source"])
    found = src.count(old)
    assert found == count, f"cell {index}: expected {count} of {old[:60]!r}, found {found}"
    src = src.replace(old, new)
    parts = src.rstrip("\n").split("\n")
    cells[index]["source"] = [p + "\n" for p in parts[:-1]] + [parts[-1]]


# --- 54: the two week ideas, told apart ------------------------------------------------
patch(54, '''def cash_week_of(date):
    """The column one payment lands in: the month it falls in, and its ISO week.

    Keyed on the month of the date rather than the ISO week's own, so a week straddling
    two months is two columns and each month's total is a real calendar month.
    """
    return f"{date.year}-{date.month:02d}-W{date.isocalendar()[1]:02d}"


def cash_week_in(period, week_key):
    """The same key for a week of a month, where the day itself is not carried."""
    return f"{period.year}-{period.month:02d}-{week_key.split('-')[-1]}"


def week_fact_key(country, cash_week):
    """A country's own copy of one week: the calendars differ, so the facts do."""
    return f"{label(country)}|{cash_week}"''',
'''def cash_week_of(date):
    """The column one payment lands in: its ISO week, and the ISO year that week belongs to.

    A week is one column and one column only. It was two for a while, with the calendar month
    on the front, so a week straddling a month end was a column under each and every month
    total was a real calendar month that tied to the budget month behind it. That is a good
    property and it was not worth what it cost: on a sheet whose columns are weeks, the same
    week number appeared twice across the top, which is the first thing anybody reading it
    asked about. The month is now a label over a run of whole weeks and nothing more, and a
    month total is the weeks under that label rather than the days in that month.

    Taken off the ISO year and not the calendar one, because a week starting in December can
    belong to the January after it - and the two halves of such a week are now one column, so
    something has to say which year holds it.
    """
    iso = date.isocalendar()
    return f"{iso[0]}-W{iso[1]:02d}"


def month_week_of(period, week_key):
    """One week *of a month*, which is not the same thing as the week cash lands in.

    A cost's share is always a share of a month, so the facts that spread it - how many of
    that month's work days this week holds, whether the month ends inside it - belong to the
    pair and not to the week alone. A straddling week therefore still has two of these, one
    per month, even though it now has only one column on the cash flow sheets.

    This is the one place the old month-on-the-front key survives, and it survives because
    here it was never about laying columns out.
    """
    return f"{period.year}-{period.month:02d}-{week_key.split('-')[-1]}"


def week_fact_key(country, month_week):
    """A country's own copy of one week of one month: the calendars differ, so the facts do."""
    return f"{label(country)}|{month_week}"''')

# --- 54: the cost rows carry both weeks ------------------------------------------------
patch(54, '''        scheduled = cash_week_in(cost.period, cost.week_key)''',
'''        # Two weeks on purpose. The month-week is what the spread facts are keyed on; the
        # plain ISO week is what a cash flow sheet joins on.
        month_week = month_week_of(cost.period, cost.week_key)''')

patch(54, '''                     scheduled, None, None,                        # and where it lands
                     week_fact_key(cost.country, scheduled),''',
'''                     cost.week_key, None, None,                    # and where it lands
                     week_fact_key(cost.country, month_week),''')

patch(54, '''            f'=IF(${days}{row}=0,${scheduled}{row},'
            f'TEXT(YEAR(${paid}{row}),"0000")&"-"&TEXT(MONTH(${paid}{row}),"00")'
            f'&"-W"&TEXT(${iso}{row},"00"))')''',
'''            f'=IF(${days}{row}=0,${scheduled}{row},'
            f'TEXT(YEAR(${thursday}{row}),"0000")&"-W"&TEXT(${iso}{row},"00"))')''')

patch(54, '''    facts = [(week_fact_key(week["country"], cash_week_in(week["period"], week["week_key"])),
              label(week["country"]),
              cash_week_in(week["period"], week["week_key"]), week["week_key"],''',
'''    facts = [(week_fact_key(week["country"], month_week_of(week["period"], week["week_key"])),
              label(week["country"]),
              month_week_of(week["period"], week["week_key"]), week["week_key"],''')

# --- 62: the COGS rows. The week key beside it is already the answer. -------------------
patch(62, '''        # And the calendar month on the front, off the payment date itself, so a week
        # straddling a month end is counted under the month whose days it is holding.
        sheet[f"{cash_week}{row}"] = (f'=TEXT(YEAR(${when}{row}),"0000")&"-"'
                                      f'&TEXT(MONTH(${when}{row}),"00")&"-W"'
                                      f'&TEXT(${week}{row},"00")')''',
'''        # The week cash lands in is now the plain ISO week, which is the cell above. Kept as
        # a column of its own rather than folded away, because two other sheets join on a
        # column called `Cash week` and the name is the interface.
        sheet[f"{cash_week}{row}"] = f"=${week_key}{row}"''')

# --- 58: the VAT settlement rows -------------------------------------------------------
patch(58, '''        sheet[f"{week}{row}"] = (f'=TEXT(YEAR(${when}{row}),"0000")&"-"'
                                 f'&TEXT(MONTH(${when}{row}),"00")&"-W"'
                                 f'&TEXT(${iso}{row},"00")')''',
'''        sheet[f"{week}{row}"] = (f'=TEXT(YEAR(${thursday}{row}),"0000")&"-W"'
                                 f'&TEXT(${iso}{row},"00")')''')

patch(58, '''        # The same three steps every other sheet in here takes to name a week: the Thursday
        # settles which week and which year it belongs to, and the month on the front comes
        # off the payment date, so a week straddling a month end counts under the month
        # whose days it is holding.''',
'''        # The same two steps every other sheet in here takes to name a week: the Thursday
        # settles which week it is and which year that week belongs to, and those are the
        # whole of the name. No month on the front - a week is one column now.''')

# --- 64: one column per ISO week, grouped under the month its Thursday falls in ---------
patch(64, '''def cash_year_weeks(year):
    """Every (month, week column) a year needs, in order.

    A week straddling two months turns up under each of them. That is what keeps a month
    column a calendar month instead of a run of whole weeks - and it matters most exactly
    where it is most likely to: a payroll on the last working day of the month.
    """
    columns = []
    for month in range(1, 13):
        first = pd.Timestamp(year=year, month=month, day=1)
        weeks = []
        for day in pd.date_range(first, first + pd.offsets.MonthEnd(0)):
            key = cash_week_of(day)
            if key not in weeks:
                weeks.append(key)
        columns.append((month, weeks))
    return columns''',
'''def cash_year_weeks(year):
    """Every ISO week of `year`, in order, grouped under the month it is labelled with.

    One column per week, and each week in exactly one group - which is the whole point. A
    week belongs to the month its Thursday falls in, the same Thursday that decides which
    ISO week a date is in and which year that week belongs to, so nothing here needs a rule
    of its own.

    The cost of it is worth saying plainly: a month total is now the weeks under that
    month's label, not the days in that month. A payroll on the 31st sits in a week that may
    be labelled with the month after, and that month's total carries it. The `Check` sheet
    still ties, because it ties on the whole horizon and every week is counted once - but a
    single month on a `CF` sheet no longer matches the same month on the `P&L` beside it.

    An ISO year is not a calendar year either. 2026 runs from Monday 29 December 2025, and
    a payment on that Monday is the first column of the 2026 sheet.
    """
    # The 4th of January is in ISO week 1 of its own year, in every year there is.
    fourth = pd.Timestamp(year=year, month=1, day=4)
    monday = fourth - pd.Timedelta(days=int(fourth.weekday()))
    columns, label_month, weeks = [], None, []
    while True:
        thursday = monday + pd.Timedelta(days=3)
        iso = thursday.isocalendar()
        if iso[0] != year:
            break
        if thursday.month != label_month:
            if label_month is not None:
                columns.append((label_month, weeks))
            label_month, weeks = thursday.month, []
        weeks.append(f"{iso[0]}-W{iso[1]:02d}")
        monday += pd.Timedelta(days=7)
    columns.append((label_month, weeks))
    return columns''')

json.dump(nb, io.open(NB, "w", encoding="utf-8", newline="\n"), indent=1, ensure_ascii=False)
print("patched cells 54, 58, 62, 64")
