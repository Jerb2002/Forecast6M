# ===== PROTOTYPE - ticket 02 - throwaway, delete once the shape is picked ==============
# Two shapes for the same paste, written side by side so they can be compared in Excel
# rather than argued about. Nothing downstream reads either sheet yet.

ST_A_SHEET, ST_A_TABLE = "Short-term A", "ShortTermA"
ST_B_SHEET, ST_B_TABLE = "Short-term B", "ShortTermB"
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

    Dated off the Thursday of the week, which is the day that decides which ISO week a date
    belongs to, and here also decides which month a straddling week is counted into. A week
    split across two months is therefore one column on the cash flow and not two, and that is
    a simplification this prototype makes rather than a rule the model already has.

    The week number is counted rather than asked for. `ISOWEEKNUM` is an Excel 2013 function,
    and a file written by openpyxl has to spell those `_xlfn.ISOWEEKNUM` or Excel opens on
    `#NAME?`. Counting whole weeks from the 1st of the Thursday's own year gives the same
    answer with functions every Excel has had for thirty years - and it is only correct
    *because* the date it counts is the Thursday.
    """
    thursday = f"({monday_cell}+3)"
    week = f'INT(({thursday}-DATE(YEAR{thursday},1,1))/7)+1'
    return (f'=YEAR{thursday}&"-"&TEXT(MONTH{thursday},"00")'
            f'&"-W"&TEXT({week},"00")')


def write_st_window(sheet, first_column, anchor):
    """The block both shapes share: one date cell, and the five weeks it opens onto.

    Returns the anchor cell, the five week-key cells one by one, and the same five as a
    range - which is everything either table needs in order to read off it.
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

    # The five weeks under it. Derived, never typed - so whichever shape wins, the keys
    # themselves come off one cell and cannot be half-updated.
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
    return anchor_cell, keys, f"${key_letter}${first}:${key_letter}${first + ST_WEEKS - 1}"


def write_st_proof(sheet, first_column, table, keys):
    """What a cash flow cell would pull out of this table, for week 1 of the window.

    The same SUMIFS the cash flow sheets already run, pointed at the paste instead of at the
    payment sheet - so the join is visible on the prototype rather than taken on trust.
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


# --- A: wide. A row per country, the five weeks across. ---------------------------------
ST_A_KEY_COLUMNS = [
    ("Country",    10, "@",      "center"),
    ("Week key",   13, "@",      "center"),
    ("Cash key",   30, "@",      "left"),
    ("Amount",     14, ST_MONEY, "right"),
]


def write_short_term_a(book):
    """Shape A: paste a block four rows deep and five columns across.

    Reads the way a human hands numbers over, and the way the short-term tool most likely
    prints them. The cost is the block on the right: the cash flow joins on one key per row,
    so a wide paste has to be unpivoted into twenty rows before anything can look it up, and
    those twenty rows are a second thing on the sheet to keep honest.

    The week columns are positions and cannot be anything else while the paste stays a
    rectangle - an Excel table will not hold a formula in its header. What week a column
    means is readable from the window block and nowhere else, which is exactly the failure
    the ticket is worried about.
    """
    paste_first = 3                                       # C, just right of the country column
    window_first = paste_first + ST_WEEKS + 1             # after a margin
    keys_first = window_first + 4
    proof_first = keys_first + len(ST_A_KEY_COLUMNS) + 1

    widths = ([14] + [13] * ST_WEEKS + [MARGIN_WIDTH] + [14, 13, 14] + [MARGIN_WIDTH]
              + [width for _, width, _, _ in ST_A_KEY_COLUMNS] + [MARGIN_WIDTH] + [12, 16])
    sheet = new_sheet(book, ST_A_SHEET, widths)
    write_title(sheet, "Short-term revenue cash - A, wide")
    write_note(sheet, "PROTOTYPE. Set the Monday in the window block, then paste five weeks "
                      "of revenue cash across each country's row. The columns are positions, "
                      "so they say nothing on their own about which weeks they hold - the "
                      "window block does. The table further right is the same numbers turned "
                      "into one row per country and week, which is the only shape the cash "
                      "flow sheets can look up.")

    anchor_cell, keys, _ = write_st_window(sheet, window_first, st_anchor_monday())

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
    write_table(sheet, ST_A_TABLE, ST_A_KEY_COLUMNS, rows, first_column=keys_first)
    country_letter = column_of(ST_A_KEY_COLUMNS, "Country", keys_first)
    week_letter = column_of(ST_A_KEY_COLUMNS, "Week key", keys_first)
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
            sheet[f"{column_of(ST_A_KEY_COLUMNS, header, keys_first)}{row}"].value = value

    write_st_proof(sheet, proof_first, ST_A_TABLE, keys)
    return sheet


# --- B: long. A row per country and week, the shape everything else here already uses. ---
ST_B_COLUMNS = [
    ("Country",    10, "@",      "center"),
    ("Week key",   13, "@",      "center"),
    ("Amount",     14, ST_MONEY, "right"),
    ("Cash key",   30, "@",      "left"),
    ("In window",  13, "@",      "center"),
]


def write_short_term_b(book):
    """Shape B: twenty rows, one per country and week, the key typed beside the amount.

    The shape every other input sheet in this workbook already has, and the one the cash flow
    reads with nothing in between. It costs the paste: twenty rows and a repeated country
    where A has a rectangle, and a key that has to arrive with the numbers.

    Because each row carries its own week, a stale paste shows up row by row rather than only
    in a date cell nobody looked at - `In window` says so against the window block.
    """
    window_first = len(ST_B_COLUMNS) + 3
    proof_first = window_first + 4

    widths = ([width for _, width, _, _ in ST_B_COLUMNS] + [MARGIN_WIDTH] + [14, 13, 14]
              + [MARGIN_WIDTH] + [12, 16])
    sheet = new_sheet(book, ST_B_SHEET, widths)
    write_title(sheet, "Short-term revenue cash - B, long")
    write_note(sheet, "PROTOTYPE. One row per country and week, which is how every other "
                      "input sheet here is built and the only shape the cash flow looks up "
                      "directly. Paste the week key beside the amount; In window says "
                      "whether that key is one of the five the window block is open on.")

    anchor_cell, keys, key_range = write_st_window(sheet, window_first, st_anchor_monday())

    # Pre-filled with the current window so the sheet reads on opening. The keys are in the
    # entry colour because in this shape they are pasted, not derived.
    monday = pd.Timestamp(st_anchor_monday())
    prefilled = []
    for country in COUNTRIES:
        for week in range(ST_WEEKS):
            thursday = monday + pd.Timedelta(days=week * 7 + 3)
            prefilled.append((label(country),
                              f"{thursday.year}-{thursday.month:02d}"
                              f"-W{thursday.isocalendar()[1]:02d}", None, None, None))
    write_table(sheet, ST_B_TABLE, ST_B_COLUMNS, prefilled, font=ENTRY)

    country_letter = column_of(ST_B_COLUMNS, "Country")
    week_letter = column_of(ST_B_COLUMNS, "Week key")
    for index in range(len(prefilled)):
        row = FIRST_DATA_ROW + index
        sheet[f"{country_letter}{row}"].font = BODY          # named by the notebook, not typed
        cell = sheet[f"{column_of(ST_B_COLUMNS, 'Cash key')}{row}"]
        cell.value = f'={country_letter}{row}&"|{ST_LINE}|"&{week_letter}{row}'
        cell.font = BODY
        window = sheet[f"{column_of(ST_B_COLUMNS, 'In window')}{row}"]
        window.value = (f'=IF(COUNTIF({key_range},{week_letter}{row})>0,"yes",'
                        f'"OUT OF WINDOW")')
        window.font = BODY

    write_st_proof(sheet, proof_first, ST_B_TABLE, keys)
    return sheet
