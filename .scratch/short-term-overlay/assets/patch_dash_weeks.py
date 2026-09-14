"""Dashboard CF by week: fifty-three week columns with the month said once over them.

String replacements on the notebook's cells, each asserted so a miss is loud rather than
silent. Idempotent: a marker in the dashboard cell says the patch is already in.
"""
import io, json, sys

NB = "forecast_dev.ipynb"
nb = json.load(io.open(NB, encoding="utf-8"))
cells = nb["cells"]


def src(i):
    return "".join(cells[i]["source"])


def put(i, text):
    parts = text.rstrip("\n").split("\n")
    cells[i]["source"] = [p + "\n" for p in parts[:-1]] + [parts[-1]]


def find(mark, kind="code"):
    hits = [i for i, c in enumerate(cells) if c["cell_type"] == kind and mark in src(i)]
    assert len(hits) == 1, (mark, hits)
    return hits[0]


def replace(i, old, new, count=1):
    s = src(i)
    assert s.count(old) == count, (i, old[:60], s.count(old))
    put(i, s.replace(old, new))


DASH = find('DASH_CF_SHEET = "Dashboard CF"')
DOC = find("### The cash dashboard", "markdown")

if "DASH_CF_WEEKS" in src(DASH):
    print("already patched")
    sys.exit(0)

# --- constants --------------------------------------------------------------------------
replace(DASH, """DASH_CF_SHEET = "Dashboard CF"
DASH_CF_TITLE = "Cash flow, month by month"
DASH_SWITCH_ROW = DASH_COMPARE_ROW + 1     # the basis dropdown, under the two selections

# The same geometry as the budget dashboard, column for column and row for row - the two
# selections, a caption, the month headers, and the lines two rows under them. So neither
# sheet has a layout of its own to learn, and the constants are the ones already declared
# over there rather than a second set saying the same thing.
#
# What differs is which rows are on it and where they are read from. There is also no basis
# row here: actual-or-budget is a fact about a P&L, and the cash runs a year past the last
# one there is a P&L for.
DASH_CF_ROW_AT = pl_row_numbers(CASH_ROWS, DASH_FIRST_ROW)
""", """DASH_CF_SHEET = "Dashboard CF"
DASH_CF_TITLE = "Cash flow, week by week"
DASH_SWITCH_ROW = DASH_COMPARE_ROW + 1     # the basis dropdown, under the two selections

# The budget dashboard's rows - the two selections, a caption, the headers, the lines two
# rows under them - over columns of its own. Cash is read a week at a time, so the block is
# fifty-three week columns wide, the most an ISO year has, with the month said once over
# the first week it labels, the way the CF sheets say it. A year with fifty-two weeks
# leaves the last column blank. The total, the comparison and the variances then sit past
# the weeks, in the order the budget dashboard keeps them.
#
# The month row is the row the budget dashboard spends on actual-or-budget, which is a fact
# about a P&L and has no meaning on cash.
DASH_CF_ROW_AT = pl_row_numbers(CASH_ROWS, DASH_FIRST_ROW)
DASH_CF_WEEKS = 53
DASH_CF_MONTH_ROW = DASH_BASIS_ROW
DASH_CF_FIRST_WEEK_COLUMN = DASH_FIRST_MONTH_COLUMN
DASH_CF_SPACER_COLUMN = DASH_CF_FIRST_WEEK_COLUMN + DASH_CF_WEEKS
DASH_CF_TOTAL_COLUMN = DASH_CF_SPACER_COLUMN + 1
DASH_CF_GAP_COLUMN = DASH_CF_TOTAL_COLUMN + 1
DASH_CF_COMPARE_COLUMN, DASH_CF_VARIANCE_COLUMN, DASH_CF_PERCENT_COLUMN = (
    DASH_CF_GAP_COLUMN + 1, DASH_CF_GAP_COLUMN + 2, DASH_CF_GAP_COLUMN + 3)
DASH_CF_WIDTHS = ([DASH_WIDTHS[0]] + [CF_WEEK_WIDTH] * DASH_CF_WEEKS
                  + DASH_WIDTHS[len(PL_MONTHS) + 1:])
""")

replace(DASH, """# A month is not in the same column from one cash flow sheet to the next. A week straddling
# two months is a column under each, and how many of those a year has is the calendar's
# business: June's total is in AO on the 2026 sheets and in AN on the 2027 ones.
#
# So the column to read is looked up rather than written in - a block at the foot of this
# sheet, hidden, one row per cash year. Its month columns sit under the same months of the
# block above, which is what makes the lookup a single INDEX down the column being written.
GEOMETRY_FIRST_ROW = CF_MATRIX_FIRST_ROW + len(CASH_ROWS) + 3
GEOMETRY_LAST_ROW = GEOMETRY_FIRST_ROW + len(CASH_YEARS) - 1
GEOMETRY_YEAR_COLUMN = DASH_LABEL_COLUMN
GEOMETRY_TOTAL_COLUMN = DASH_TOTAL_COLUMN
""", """# A week is not in the same column from one cash flow sheet to the next. Every month has a
# total column after its weeks, and which month a week is labelled with is the calendar's
# business: W23 is in AD on the 2026 sheets and in AC on the 2027 ones. Nor is the month a
# week sits under the same from year to year.
#
# So both are looked up rather than written in - a block at the foot of this sheet, hidden,
# two rows per cash year: the column each week is in, and the month said over the first
# week it labels. Its week columns sit under the same weeks of the block above, which is
# what makes each lookup a single INDEX down the column being written. A year of fifty-two
# weeks leaves its fifty-third cell blank, and the block above reads that as nothing to
# show.
GEOMETRY_FIRST_ROW = CF_MATRIX_FIRST_ROW + len(CASH_ROWS) + 3
GEOMETRY_LAST_ROW = GEOMETRY_FIRST_ROW + len(CASH_YEARS) - 1
GEOMETRY_MONTH_FIRST_ROW = GEOMETRY_LAST_ROW + 1
GEOMETRY_MONTH_LAST_ROW = GEOMETRY_MONTH_FIRST_ROW + len(CASH_YEARS) - 1
GEOMETRY_YEAR_COLUMN = DASH_LABEL_COLUMN
GEOMETRY_TOTAL_COLUMN = DASH_CF_TOTAL_COLUMN
""")

# --- the geometry ------------------------------------------------------------------------
replace(DASH, """    return columns, position           # the twelve months, and the year beside them


def write_cf_geometry(sheet):
    \"\"\"One hidden row per cash year: which column each month's total is in, and the year's.

    The only figures on the sheet rather than formulas, and they are here because they are
    the one thing Excel cannot work out for itself - the shape of a calendar on a sheet it
    is not looking at.
    \"\"\"
    for offset, year in enumerate(CASH_YEARS):
        row = GEOMETRY_FIRST_ROW + offset
        sheet.row_dimensions[row].hidden = True
        sheet.cell(row=row, column=GEOMETRY_YEAR_COLUMN, value=year)
        months, year_column = cf_month_columns(year)
        for index, column in enumerate(months):
            sheet.cell(row=row, column=DASH_FIRST_MONTH_COLUMN + index, value=column)
        sheet.cell(row=row, column=GEOMETRY_TOTAL_COLUMN, value=year_column)


def geometry_lookup(year_ref, column):
    \"\"\"The column an INDIRECT below should read, off the block at the foot of the sheet.

    `column` is the column being written, and the block is laid out under it - so a month
    finds its own column, and the total column finds the year.

    VALUE around the year because a dropdown hands back what it was given a list of, and a
    year matched as text against a year held as a number finds nothing.
    \"\"\"
    year_letter = get_column_letter(GEOMETRY_YEAR_COLUMN)
    letter = get_column_letter(column)
    years = f"${year_letter}${GEOMETRY_FIRST_ROW}:${year_letter}${GEOMETRY_LAST_ROW}"
    block = f"${letter}${GEOMETRY_FIRST_ROW}:${letter}${GEOMETRY_LAST_ROW}"
    return f"INDEX({block},MATCH(VALUE({year_ref}),{years},0))"
""", """    return columns, position           # the twelve months, and the year beside them


def cf_week_columns(year):
    \"\"\"Where each week lands on a `CF` sheet for `year`, with the month said over the first
    week it labels, and the year column after them.

    Laid out the same way write_cash_flow() lays it out - a column per week of a month and
    one more for that month's total - rather than read back off the finished sheet, so the
    two cannot drift apart.
    \"\"\"
    weeks, position = [], CF_FIRST_DATA_COLUMN
    for month, keys in cash_year_weeks(year):
        for index, _ in enumerate(keys):
            weeks.append((position, PL_MONTHS[month - 1] if index == 0 else ""))
            position += 1
        position += 1                  # that month's total
    return weeks, position             # the weeks, and the year past the last total


def write_cf_geometry(sheet):
    \"\"\"Two hidden rows per cash year: which column each week is in, and the month over it;
    and in the total column, which column the year is in.

    The only figures on the sheet rather than formulas, and they are here because they are
    the one thing Excel cannot work out for itself - the shape of a calendar on a sheet it
    is not looking at.
    \"\"\"
    for offset, year in enumerate(CASH_YEARS):
        row, month_row = GEOMETRY_FIRST_ROW + offset, GEOMETRY_MONTH_FIRST_ROW + offset
        for each in (row, month_row):
            sheet.row_dimensions[each].hidden = True
            sheet.cell(row=each, column=GEOMETRY_YEAR_COLUMN, value=year)
        weeks, year_column = cf_week_columns(year)
        for index, (column, month) in enumerate(weeks):
            sheet.cell(row=row, column=DASH_CF_FIRST_WEEK_COLUMN + index, value=column)
            if month:
                sheet.cell(row=month_row, column=DASH_CF_FIRST_WEEK_COLUMN + index,
                           value=month)
        sheet.cell(row=row, column=GEOMETRY_TOTAL_COLUMN, value=year_column)


def geometry_lookup(year_ref, column, months=False):
    \"\"\"The column an INDIRECT below should read, off the block at the foot of the sheet -
    or, with `months`, the month label over it.

    `column` is the column being written, and the block is laid out under it - so a week
    finds its own column, and the total column finds the year. A reference, not a value,
    so ISBLANK on it says whether this year has this week at all.

    VALUE around the year because a dropdown hands back what it was given a list of, and a
    year matched as text against a year held as a number finds nothing.
    \"\"\"
    year_letter = get_column_letter(GEOMETRY_YEAR_COLUMN)
    letter = get_column_letter(column)
    first, last = ((GEOMETRY_MONTH_FIRST_ROW, GEOMETRY_MONTH_LAST_ROW) if months
                   else (GEOMETRY_FIRST_ROW, GEOMETRY_LAST_ROW))
    years = f"${year_letter}${first}:${year_letter}${last}"
    block = f"${letter}${first}:${letter}${last}"
    return f"INDEX({block},MATCH(VALUE({year_ref}),{years},0))"
""")

# --- the sheet ---------------------------------------------------------------------------
replace(DASH, """    sheet = new_sheet(book, DASH_CF_SHEET, DASH_WIDTHS)
    write_title(sheet, DASH_CF_SHEET)
""", """    sheet = new_sheet(book, DASH_CF_SHEET, DASH_CF_WIDTHS)
    write_title(sheet, DASH_CF_SHEET)
""")

replace(DASH, """    # --- the headers
    headers = [(DASH_LABEL_COLUMN, PL_UNIT)] + [
        (DASH_FIRST_MONTH_COLUMN + offset, month) for offset, month in enumerate(PL_MONTHS)]
    headers += [(DASH_TOTAL_COLUMN, "Total"), (DASH_COMPARE_COLUMN, None),
                (DASH_VARIANCE_COLUMN, "Variance"), (DASH_PERCENT_COLUMN, "Var %")]
    for column, text in headers:
        cell = sheet.cell(row=DASH_HEADER_ROW, column=column, value=text)
        cell.font, cell.border = HEADER, UNDERLINE
        cell.alignment = Alignment(horizontal="left" if column == DASH_LABEL_COLUMN
                                   else "right")
    # The comparison column names itself, so it is obvious which way the variance runs.
    sheet.cell(row=DASH_HEADER_ROW, column=DASH_COMPARE_COLUMN,
               value=f'={COMPARE_COUNTRY_CELL}&" "&{COMPARE_YEAR_CELL}')

    # --- the lines
    write_cf_geometry(sheet)
    write_dash_cash_block(sheet)
    for column in (DASH_SPACER_COLUMN, DASH_GAP_COLUMN):
        sheet.column_dimensions[get_column_letter(column)].hidden = True
    sheet.freeze_panes = f"{get_column_letter(DASH_FIRST_MONTH_COLUMN)}{DASH_FIRST_ROW}"
""", """    # --- the headers: the month once over the first week it labels, then the weeks. Both
    # off the geometry block, because which weeks a month has is the year's business; a
    # week the year does not have (a fifty-third, most years) reads blank in both rows.
    for offset in range(DASH_CF_WEEKS):
        column = DASH_CF_FIRST_WEEK_COLUMN + offset
        month = sheet.cell(row=DASH_CF_MONTH_ROW, column=column,
                           value=f'={geometry_lookup(YEAR_CELL, column, months=True)}&""')
        month.font, month.alignment = HEADER, Alignment(horizontal="left")
        week = sheet.cell(row=DASH_HEADER_ROW, column=column,
                          value=f'=IF(ISBLANK({geometry_lookup(YEAR_CELL, column)}),"",'
                                f'"W{offset + 1:02d}")')
        week.font, week.border = HEADER, UNDERLINE
        week.alignment = Alignment(horizontal="right")
    headers = [(DASH_LABEL_COLUMN, PL_UNIT), (DASH_CF_TOTAL_COLUMN, "Total"),
               (DASH_CF_COMPARE_COLUMN, None), (DASH_CF_VARIANCE_COLUMN, "Variance"),
               (DASH_CF_PERCENT_COLUMN, "Var %")]
    for column, text in headers:
        cell = sheet.cell(row=DASH_HEADER_ROW, column=column, value=text)
        cell.font, cell.border = HEADER, UNDERLINE
        cell.alignment = Alignment(horizontal="left" if column == DASH_LABEL_COLUMN
                                   else "right")
    # The comparison column names itself, so it is obvious which way the variance runs.
    sheet.cell(row=DASH_HEADER_ROW, column=DASH_CF_COMPARE_COLUMN,
               value=f'={COMPARE_COUNTRY_CELL}&" "&{COMPARE_YEAR_CELL}')

    # --- the lines
    write_cf_geometry(sheet)
    write_dash_cash_block(sheet)
    for column in (DASH_CF_SPACER_COLUMN, DASH_CF_GAP_COLUMN):
        sheet.column_dimensions[get_column_letter(column)].hidden = True
    sheet.freeze_panes = f"{get_column_letter(DASH_CF_FIRST_WEEK_COLUMN)}{DASH_FIRST_ROW}"
""")

replace(DASH, """def write_dash_cash_block(sheet):
    \"\"\"The twelve months, the year, and the comparison beside it, line by line.\"\"\"
    total_letter = get_column_letter(DASH_TOTAL_COLUMN)
    compare_letter = get_column_letter(DASH_COMPARE_COLUMN)
    variance_letter = get_column_letter(DASH_VARIANCE_COLUMN)
    first = get_column_letter(DASH_FIRST_MONTH_COLUMN)
    last = get_column_letter(DASH_FIRST_MONTH_COLUMN + len(PL_MONTHS) - 1)
""", """def write_dash_cash_block(sheet):
    \"\"\"The weeks, the year, and the comparison beside it, line by line.\"\"\"
    total_letter = get_column_letter(DASH_CF_TOTAL_COLUMN)
    compare_letter = get_column_letter(DASH_CF_COMPARE_COLUMN)
    variance_letter = get_column_letter(DASH_CF_VARIANCE_COLUMN)
    first = get_column_letter(DASH_CF_FIRST_WEEK_COLUMN)
    last = get_column_letter(DASH_CF_FIRST_WEEK_COLUMN + DASH_CF_WEEKS - 1)
""")

replace(DASH, """        for offset in range(len(PL_MONTHS)):
            column = DASH_FIRST_MONTH_COLUMN + offset
            write_cf_cell(sheet, row, column, kind, "=" + pl_lookup(
                COUNTRY_CELL, YEAR_CELL, CASH_ROW_AT[key],
                geometry_lookup(YEAR_CELL, column), prefix=CF_SHEET_PREFIX))

        # The year. Summed from the twelve months rather than taken off the cash flow
        # sheet's own year column, so it cannot disagree with the figures beside it.
        write_cf_cell(sheet, row, DASH_TOTAL_COLUMN, kind,
                      f"=SUM({first}{row}:{last}{row})")

        write_cf_cell(sheet, row, DASH_COMPARE_COLUMN, kind, "=" + pl_lookup(
            COMPARE_COUNTRY_CELL, COMPARE_YEAR_CELL, CASH_ROW_AT[key],
            geometry_lookup(COMPARE_YEAR_CELL, GEOMETRY_TOTAL_COLUMN),
            prefix=CF_SHEET_PREFIX))
        write_cf_cell(sheet, row, DASH_VARIANCE_COLUMN, kind,
                      f"={total_letter}{row}-{compare_letter}{row}")
        percent = write_cf_cell(
            sheet, row, DASH_PERCENT_COLUMN, kind,
            f'=IFERROR({variance_letter}{row}/ABS({compare_letter}{row}),"")')
""", """        for offset in range(DASH_CF_WEEKS):
            column = DASH_CF_FIRST_WEEK_COLUMN + offset
            where = geometry_lookup(YEAR_CELL, column)
            # Blank, not nought, for a week the year does not have: SUM walks past text.
            write_cf_cell(sheet, row, column, kind, f'=IF(ISBLANK({where}),"",' + pl_lookup(
                COUNTRY_CELL, YEAR_CELL, CASH_ROW_AT[key], where,
                prefix=CF_SHEET_PREFIX) + ")")

        # The year. Summed from the weeks rather than taken off the cash flow sheet's own
        # year column, so it cannot disagree with the figures beside it.
        write_cf_cell(sheet, row, DASH_CF_TOTAL_COLUMN, kind,
                      f"=SUM({first}{row}:{last}{row})")

        write_cf_cell(sheet, row, DASH_CF_COMPARE_COLUMN, kind, "=" + pl_lookup(
            COMPARE_COUNTRY_CELL, COMPARE_YEAR_CELL, CASH_ROW_AT[key],
            geometry_lookup(COMPARE_YEAR_CELL, GEOMETRY_TOTAL_COLUMN),
            prefix=CF_SHEET_PREFIX))
        write_cf_cell(sheet, row, DASH_CF_VARIANCE_COLUMN, kind,
                      f"={total_letter}{row}-{compare_letter}{row}")
        percent = write_cf_cell(
            sheet, row, DASH_CF_PERCENT_COLUMN, kind,
            f'=IFERROR({variance_letter}{row}/ABS({compare_letter}{row}),"")')
""")

# --- the write-up -----------------------------------------------------------------------
replace(DOC, """The same sheet as `Dashboard Budget`, over the cash flow sheets instead of the P&Ls. Two dropdowns
pick a country and a year, two more pick what to compare it with, and the block under them
is that country's cash flow by month, with the twelve months, the year, the comparison, the
variance and the variance per cent. The matrix below it is every country in its own column
for the selected year, and the group beside them. Nothing on it is typed except the four
dropdowns, and everything on it is a formula reading a `CF` sheet - so it moves when they
move.
""", """The same sheet as `Dashboard Budget`, over the cash flow sheets instead of the P&Ls. Two dropdowns
pick a country and a year, two more pick what to compare it with, and the block under them
is that country's cash flow by week: fifty-three week columns with the month said once over
the first week it labels, then the year, the comparison, the variance and the variance per
cent. A year of fifty-two weeks leaves the last column blank. The matrix below it is every
country in its own column for the selected year, and the group beside them. Nothing on it
is typed except the five dropdowns, and everything on it is a formula reading a `CF` sheet
- so it moves when they move, the basis included.
""")

replace(DOC, """**The one thing that is genuinely harder than on the budget dashboard** is which column to
read. A P&L's January is always column C. A cash flow sheet's January is a run of week
columns and then a total, and how many weeks a month is labelled with depends on the
calendar - four in most months, five in some - so a month's total column moves from year to
year.

So the column is looked up rather than written in. A block at the foot of the sheet, hidden,
holds one row per cash year: the column each month's total is in, and the column the year is
in. Its months sit under the same months of the block above, so the lookup is one `INDEX`
straight down the column being written, and a figure here stays a single `INDIRECT` instead
of one per year nested inside an `IF`.
""", """**The one thing that is genuinely harder than on the budget dashboard** is which column to
read. A P&L's January is always column C. A cash flow sheet's week 23 is somewhere in the
twenties, after however many month totals the calendar has put in front of it, and which
month it sits under moves from year to year too.

So both are looked up rather than written in. A block at the foot of the sheet, hidden,
holds two rows per cash year: the column each week is in, and the month over the first week
it labels; the year's column sits in the total column. Its weeks sit under the same weeks of
the block above, so each lookup is one `INDEX` straight down the column being written, and
a figure here stays a single `INDIRECT` instead of one per year nested inside an `IF`. The
block is written by the same function that lays the `CF` sheets out, so the two cannot
drift apart.
""")

json.dump(nb, io.open(NB, "w", encoding="utf-8"), indent=1, ensure_ascii=False)
io.open(NB, "a", encoding="utf-8").write("\n")
print("patched")
