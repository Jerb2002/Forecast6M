# ===== The short-term paste - ticket 02 ================================================
# The shape was picked by building both in Excel and looking at them: wide won, because a
# CFO reads a rectangle and does not read a twenty-row table. Nothing reads this sheet yet;
# ticket 05 wires it into the cash flow.

ST_SHEET, ST_TABLE = "Short-term input", "ShortTerm"
ST_WEEKS = 5                                   # how many weeks the short-term tool covers
ST_LINE = STREAM_LABELS["revenue"]             # the cash flow line an override would replace
ST_MONEY = "#,##0"                             # whole units, like Cash amount


def st_anchor_monday():
    """The Monday the window opens on, defaulted to the Monday of the week this is built in."""
    today = pd.Timestamp.today().normalize()
    return (today - pd.Timedelta(days=int(today.weekday()))).to_pydatetime()


def st_week_key_formula(monday_cell):
    """The `Cash week` a pasted week lands in, worked out in the workbook.

    Digits only - no `yyyy` or `mm` inside a TEXT format - because those letters change with
    the language Excel runs in, and the key would come out spelt differently on a Danish
    machine than on an English one.

    Dated off the Thursday of the week, which is the day that decides both which ISO week a
    date is in and which year that week belongs to. A pasted week is one column on the cash
    flow, the same as every other week there - no month on the front, so a week straddling a
    month end needs no rule and gets no second copy of itself.

    The week number is counted rather than asked for. `ISOWEEKNUM` is an Excel 2013 function,
    and a file written by openpyxl has to spell those `_xlfn.ISOWEEKNUM` or Excel opens on
    `#NAME?`. Counting whole weeks from the 1st of the Thursday's own year gives the same
    answer with functions every Excel has had for thirty years - and it is only correct
    *because* the date it counts is the Thursday.
    """
    thursday = f"({monday_cell}+3)"
    week = f'INT(({thursday}-DATE(YEAR{thursday},1,1))/7)+1'
    return f'=YEAR{thursday}&"-W"&TEXT({week},"00")'


def write_st_window(sheet, first_column, anchor):
    """One date cell, and the five weeks it opens onto.

    The Monday is the only thing typed here. Everything else on the sheet is worked out from
    it, so the five weeks cannot be half-updated - which is the one thing about the window
    that is *not* left to whoever is pasting.

    Returns the anchor cell and the five week-key cells, which is what the table reads.
    """
    monday_letter = get_column_letter(first_column + 1)
    key_letter = get_column_letter(first_column + 2)

    head = sheet.cell(row=HEADER_ROW, column=first_column, value="The window")
    head.font, head.border = HEADER, UNDERLINE
    for offset in (1, 2):
        sheet.cell(row=HEADER_ROW, column=first_column + offset).border = UNDERLINE

    anchor_cell = f"{monday_letter}{HEADER_ROW + 1}"
    stale = f"INT((TODAY()-{anchor_cell})/7)"
    rows = [("As of Monday", anchor, "yyyy-mm-dd", ENTRY),
            ("Today", "=TODAY()", "yyyy-mm-dd", BODY),
            ("Weeks old", f"={stale}", "0", BODY),
            ("Paste is", f'=IF({anchor_cell}>TODAY(),"dated ahead - check it",'
                         f'IF({stale}=0,"this week",'
                         f'IF({stale}=1,"last week - repaste",'
                         f'{stale}&" weeks old - repaste")))', "@", BODY)]
    for offset, (text, value, number_format, font) in enumerate(rows):
        sheet.cell(row=HEADER_ROW + 1 + offset, column=first_column, value=text).font = BODY
        cell = sheet.cell(row=HEADER_ROW + 1 + offset, column=first_column + 1, value=value)
        cell.font, cell.number_format = font, number_format
        cell.alignment = Alignment(horizontal="center")

    # The five weeks under it, derived and never typed.
    first = HEADER_ROW + len(rows) + 3
    for offset, text in enumerate(("Position", "Monday", "Week key")):
        cell = sheet.cell(row=first - 1, column=first_column + offset, value=text)
        cell.font, cell.border = HEADER, UNDERLINE
    keys = []
    for week in range(ST_WEEKS):
        row = first + week
        sheet.cell(row=row, column=first_column, value=f"Week {week + 1}").font = BODY
        monday = sheet.cell(row=row, column=first_column + 1,
                            value=f"={anchor_cell}+{week * 7}")
        monday.font, monday.number_format = BODY, "yyyy-mm-dd"
        monday.alignment = Alignment(horizontal="center")
        key = sheet.cell(row=row, column=first_column + 2,
                         value=st_week_key_formula(f"{monday_letter}{row}"))
        key.font, key.number_format = BODY, "@"
        key.alignment = Alignment(horizontal="center")
        keys.append(f"${key_letter}${row}")
    return anchor_cell, keys


def write_st_proof(sheet, first_column, table, keys):
    """What a cash flow cell would pull out of this table, for week 1 of the window.

    The same SUMIFS the cash flow sheets already run, pointed at the paste instead of at the
    payment sheet - so the join is visible on the sheet rather than taken on trust. It is
    here because nothing downstream reads this sheet yet; ticket 05 is what makes it real.
    """
    for offset, text in enumerate(("Country", "CF reads, week 1")):
        cell = sheet.cell(row=HEADER_ROW, column=first_column + offset, value=text)
        cell.font, cell.border = HEADER, UNDERLINE
        cell.alignment = Alignment(horizontal="center" if offset else "left")
    code_letter = get_column_letter(first_column)
    for offset, country in enumerate(COUNTRIES):
        row = FIRST_DATA_ROW + offset
        sheet.cell(row=row, column=first_column, value=label(country)).font = BODY
        cell = sheet.cell(row=row, column=first_column + 1,
                          value=f'=SUMIFS({table}[Amount],{table}[Cash key],'
                                f'${code_letter}{row}&"|{ST_LINE}|"&{keys[0]})')
        cell.font, cell.number_format = BODY, ST_MONEY
        cell.alignment = Alignment(horizontal="right")


# --- the unpivot, which is what the wide shape costs ------------------------------------
ST_KEY_COLUMNS = [
    ("Country",    10, "@",      "center"),
    ("Week key",   13, "@",      "center"),
    ("Cash key",   30, "@",      "left"),
    ("Amount",     14, ST_MONEY, "right"),
]


def write_short_term(book):
    """Five weeks of revenue cash: a block four rows deep and five columns across.

    Wide, because a CFO reads a rectangle and does not read a twenty-row table. Both shapes
    were built and opened in Excel before this one was kept.

    The cost is the block further right: the cash flow joins on one key per row, so the
    rectangle is unpivoted into twenty rows before anything can look it up. Those twenty rows
    are formulas back into the paste, so they cannot drift from it.

    The week columns do label themselves - the paste block is a plain styled range rather
    than an Excel table, and a table is the only thing here that could not hold a formula in
    its header. What they cannot do is label the *amounts*. Move the Monday without re-pasting
    and last week's figures silently become this week's, because a figure in this shape knows
    only which column it is in. The window block says how many weeks old the Monday is, and
    that is the whole of the guard: it catches a Monday nobody touched, and cannot catch a
    Monday touched on its own. See ticket 04 for what a CFO does on a Monday, in order.
    """
    paste_first = 3                                       # C, just right of the country column
    window_first = paste_first + ST_WEEKS + 1             # after a margin
    keys_first = window_first + 4
    proof_first = keys_first + len(ST_KEY_COLUMNS) + 1

    widths = ([14] + [13] * ST_WEEKS + [MARGIN_WIDTH] + [14, 13, 14] + [MARGIN_WIDTH]
              + [width for _, width, _, _ in ST_KEY_COLUMNS] + [MARGIN_WIDTH] + [12, 16])
    sheet = new_sheet(book, ST_SHEET, widths)
    write_title(sheet, "Short-term revenue cash, five weeks")
    write_note(sheet, "Every Monday: set the Monday in the window block first, then paste "
                      "five weeks of revenue cash across each country's row. Do both, in "
                      "that order - the figures carry no date of their own, so moving the "
                      "Monday and leaving them alone shifts them a week without saying so. "
                      "The window block says how old the Monday is. The table further right "
                      "is the same figures with one row per country and week, which is the "
                      "only shape the cash flow sheets can look up; it is not typed.")

    anchor_cell, keys = write_st_window(sheet, window_first, st_anchor_monday())

    # --- the paste block. Not a named table: an Excel table cannot hold a formula in its
    # header, and these headers have to show the live week or they show nothing at all.
    head = sheet.cell(row=HEADER_ROW, column=2, value="Country")
    head.font, head.border = HEADER, UNDERLINE
    for week in range(ST_WEEKS):
        cell = sheet.cell(row=HEADER_ROW, column=paste_first + week, value=f"={keys[week]}")
        cell.font, cell.border, cell.number_format = HEADER, UNDERLINE, "@"
        cell.alignment = Alignment(horizontal="center")
    for offset, country in enumerate(COUNTRIES):
        row = FIRST_DATA_ROW + offset
        sheet.cell(row=row, column=2, value=label(country)).font = BODY
        for week in range(ST_WEEKS):
            cell = sheet.cell(row=row, column=paste_first + week)
            cell.font, cell.number_format = ENTRY, ST_MONEY
            cell.alignment = Alignment(horizontal="right")
    sheet.freeze_panes = f"{get_column_letter(paste_first)}{FIRST_DATA_ROW}"

    # --- the unpivot the wide shape costs: twenty rows, every cell a formula back into it.
    rows = [(None, None, None, None) for _ in COUNTRIES for _ in range(ST_WEEKS)]
    write_table(sheet, ST_TABLE, ST_KEY_COLUMNS, rows, first_column=keys_first)
    country_letter = column_of(ST_KEY_COLUMNS, "Country", keys_first)
    week_letter = column_of(ST_KEY_COLUMNS, "Week key", keys_first)
    for index in range(len(COUNTRIES) * ST_WEEKS):
        row = FIRST_DATA_ROW + index
        country, week = divmod(index, ST_WEEKS)
        paste_row = FIRST_DATA_ROW + country
        paste_letter = get_column_letter(paste_first + week)
        for header, value in (
                ("Country",  f"=$B{paste_row}"),
                ("Week key", f"={keys[week]}"),
                ("Cash key", f'={country_letter}{row}&"|{ST_LINE}|"&{week_letter}{row}'),
                ("Amount",   f"={paste_letter}{paste_row}")):
            sheet[f"{column_of(ST_KEY_COLUMNS, header, keys_first)}{row}"].value = value

    write_st_proof(sheet, proof_first, ST_TABLE, keys)
    return sheet
