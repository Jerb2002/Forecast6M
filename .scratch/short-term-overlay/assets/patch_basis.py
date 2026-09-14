"""Ticket 05: wire the basis switch into forecast_dev.ipynb.

String replacements on the notebook's cells, each asserted so a miss is loud rather than
silent. Idempotent: a marker in the CF cell says the patch is already in.
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


def find(mark):
    hits = [i for i, c in enumerate(cells) if c["cell_type"] == "code" and mark in src(i)]
    assert len(hits) == 1, (mark, hits)
    return hits[0]


def replace(i, old, new, count=1):
    s = src(i)
    assert s.count(old) == count, (i, old[:60], s.count(old))
    put(i, s.replace(old, new))


IMPORTS = find("from openpyxl import Workbook")
CF = find('CF_SHEET_PREFIX = "CF"')
DASH = find('DASH_CF_SHEET = "Dashboard CF"')
ST = find('ST_SHEET, ST_TABLE = "Short-term input", "ShortTerm"')

if "CF_BASIS_ROWS" in src(CF):
    print("already patched")
    sys.exit(0)

# --- imports ---------------------------------------------------------------------------
replace(IMPORTS,
        "from openpyxl.utils import get_column_letter\n",
        "from openpyxl.utils import get_column_letter\n"
        "from openpyxl.workbook.defined_name import DefinedName\n")

# --- the CF sheet: constants ------------------------------------------------------------
replace(CF, """CF_YEAR_ROW, CF_MONTH_ROW, CF_HEADER_ROW, CF_KEY_ROW = 4, 5, 6, 7
CF_FIRST_ROW = CF_KEY_ROW + 2
""", """CF_YEAR_ROW, CF_MONTH_ROW, CF_HEADER_ROW, CF_KEY_ROW = 4, 5, 6, 7

# --- the basis ----------------------------------------------------------------------------
# One dropdown on Dashboard CF, read on every CF sheet through the workbook name `Basis`.
# Each sheet keeps its own copy in the note cell under the title, so the figures say which
# basis they are on - and so one sheet can be pinned to another basis by typing over it.
BASIS_NAME, FIRST_PASTE_NAME = "Basis", "FirstPasteWeek"
BASIS_BUDGET, BASIS_OVERRIDE, BASIS_DIFFERENCE = ("Budget only", "Short-term override",
                                                  "Difference")
BASES = (BASIS_BUDGET, BASIS_OVERRIDE, BASIS_DIFFERENCE)
CF_BASIS_CELL = f"$B${NOTE_ROW}"           # this sheet's basis: =Basis unless typed over
CF_DIFF_CELL = f"$C${NOTE_ROW}"            # 1 under Difference; hidden column
CF_RATE_CELL = f"$D${NOTE_ROW}"            # this country's VAT rate; hidden column

# How the switch reaches the figures. "key": through hidden rows under the key row, one
# figure per week column. The rows the paste never touches look up a *key* off those rows -
# the real week when the column is live on this basis, and a key that matches nothing when
# it is not - so their formulas stay the one lookup they always were, and only the three
# rows the paste replaces carry an IF. "if": every cell wrapped in an IF instead, which is
# the obvious way and the one the ticket suspected of being slow. Both were built and timed
# in Excel; see ticket 05.
CF_SWITCH = "key"

CF_BASIS_ROWS = [
    ("live",        "Live on this basis"),          # 1 shown; 0 before the paste window
    ("override",    "Overridden by the paste"),     # 1 in one of the five pasted weeks
    ("mult",        "Budget figure survives"),      # live, and not Difference
    ("plain_key",   "Key for the untouched rows"),  # the week, or "-" to find nothing
    ("paste",       "Pasted revenue cash"),
    ("ratio",       "Budget duty over revenue"),    # r, for the month this week sits under
    ("service",     "Service part of the paste"),   # paste / (1 + r + v)
    ("rev_budget",  "Revenue, budget"),
    ("rev_direct",  "Revenue, direct leg only"),
    ("duty_budget", "Revenue duty, budget"),
    ("duty_direct", "Revenue duty, direct leg only"),
    ("vat_budget",  "VAT on sales, budget"),
    ("vat_direct",  "VAT on sales, direct leg only"),
]
CF_BASIS_FIRST_ROW = CF_KEY_ROW + 1
CF_BASIS_ROW_AT = {key: CF_BASIS_FIRST_ROW + offset
                   for offset, (key, _) in enumerate(CF_BASIS_ROWS)}
CF_FIRST_ROW = CF_BASIS_FIRST_ROW + len(CF_BASIS_ROWS) + 1
# The three rows the paste replaces, and which helper rows each one is built from.
CF_OVERRIDDEN = {"in_revenue": "rev", "in_duty": "duty", "in_vat": "vat"}
""")

# --- the CF sheet: the helper rows and the override formulas ------------------------------
replace(CF, '''def write_cash_flow(book, country, year):
    """One country's cash for one year: its weeks across, grouped under their months."""
''', '''def write_cf_basis_rows(sheet, code, year, months, weeks):
    """The hidden rows the basis reaches the figures through, one figure per week column.

    `ShortTerm` and the budget table are both small, so the lookups here are cheap; what
    costs is the six budget-basis rows, which are the same lookups the three overridden
    lines used to carry themselves, with a `Direct`-leg copy of each beside it. That is the
    price of ticket 03: an overridden week keeps the budget's factored legs and drops only
    the leg the customer pays, so the sheet has to be able to tell the two apart.
    """
    at = CF_BASIS_ROW_AT
    line_letter = get_column_letter(CF_LINE_COLUMN)
    share_letter = get_column_letter(CF_VAT_COLUMN)
    rev_row, duty_row = CASH_ROW_AT["in_revenue"], CASH_ROW_AT["in_duty"]
    v = f"${share_letter}${rev_row}*{CF_RATE_CELL}"      # the VAT rate on the service part
    direct = f'{PAYMENT_TABLE}[Leg],"{LEG_LABELS["direct"]}"'

    for key, text in CF_BASIS_ROWS:
        sheet.row_dimensions[at[key]].hidden = True
        sheet.cell(row=at[key], column=CF_LABEL_COLUMN, value=text).font = NOTE

    for month, _ in months:
        budget_key = f"{year}-{month:02d}-{code}"
        budget = (f'SUMIFS({BUDGET_TABLE}[Amount],{BUDGET_TABLE}[Key],"{budget_key}",'
                  f'{BUDGET_TABLE}[Stream],"{{}}")')
        for position, _ in weeks[month]:
            letter = get_column_letter(position)
            week = f"{letter}${CF_KEY_ROW}"

            def ref(key):
                return f"{letter}${at[key]}"

            def traded(row, leg=""):
                return (f'SUMIFS({PAYMENT_TABLE}[Cash amount],{PAYMENT_TABLE}[Cash key],'
                        f'"{code}|"&${line_letter}${row}&"|"&{week}{leg})')

            def vat(leg=""):
                # The same by-stream lookup the VAT on sales row has always run.
                terms = [f'SUMIFS({PAYMENT_TABLE}[Cash amount],'
                         f'{PAYMENT_TABLE}[Country],"{code}",'
                         f'{PAYMENT_TABLE}[Stream],${line_letter}${row},'
                         f'{PAYMENT_TABLE}[Cash week],{week}{leg})'
                         f'*${share_letter}${row}'
                         for row in (rev_row, duty_row)]
                return f"({'+'.join(terms)})*{CF_RATE_CELL}"

            formulas = {
                # ST_TABLE and ST_LINE are named on the short-term sheet, further down;
                # nothing here runs before the build cell, by which point they exist.
                "live": f'=IF(OR({CF_BASIS_CELL}="{BASIS_BUDGET}",'
                        f'{week}>={FIRST_PASTE_NAME}),1,0)',
                "override": f'=IF(AND({CF_BASIS_CELL}<>"{BASIS_BUDGET}",'
                            f'COUNTIF({ST_TABLE}[Week key],{week})>0),1,0)',
                "mult": f'={ref("live")}*(1-{CF_DIFF_CELL})',
                "plain_key": f'=IF({ref("mult")}=1,{week},"-")',
                "paste": f'=SUMIFS({ST_TABLE}[Amount],{ST_TABLE}[Cash key],'
                         f'"{code}|{ST_LINE}|"&{week})',
                "ratio": f'=IFERROR({budget.format(STREAM_LABELS["revenue_duty"])}'
                         f'/{budget.format(STREAM_LABELS["revenue"])},0)',
                "service": f'={ref("paste")}/(1+{ref("ratio")}+{v})',
                "rev_budget": "=" + traded(rev_row),
                "rev_direct": "=" + traded(rev_row, "," + direct),
                "duty_budget": "=" + traded(duty_row),
                "duty_direct": "=" + traded(duty_row, "," + direct),
                "vat_budget": "=" + vat(),
                "vat_direct": "=" + vat("," + direct),
            }
            for key, formula in formulas.items():
                sheet.cell(row=at[key], column=position, value=formula)


def cf_override_formula(key, letter):
    """One of the three replaced rows, in one week column.

    In an overridden week: the paste's part of this row, plus the budget's factored legs
    (the budget figure less its direct leg). Under Difference the budget figure comes off
    again, so what is left is the paste's part less the direct leg it stands in for. In
    any other week: the budget figure, or nothing, as the basis says.
    """
    at, stem = CF_BASIS_ROW_AT, CF_OVERRIDDEN[key]

    def ref(name):
        return f"{letter}${at[name]}"

    rev_row = CASH_ROW_AT["in_revenue"]
    share = f"${get_column_letter(CF_VAT_COLUMN)}${rev_row}"
    part = {"rev": ref("service"),
            "duty": f'{ref("service")}*{ref("ratio")}',
            "vat": f'{ref("service")}*{share}*{CF_RATE_CELL}'}[stem]
    return (f'=IF({ref("override")}=1,'
            f'{part}+{ref(stem + "_budget")}*(1-{CF_DIFF_CELL})-{ref(stem + "_direct")},'
            f'{ref(stem + "_budget")}*{ref("mult")})')


def write_cash_flow(book, country, year):
    """One country's cash for one year: its weeks across, grouped under their months."""
''')

# The basis cells under the title, and the country rate moved up so they can share it.
replace(CF, '''    write_title(sheet, f"{CF_TITLE} - {code} {year}")
    for column in (CF_LINE_COLUMN, CF_VAT_COLUMN):
        sheet.column_dimensions[get_column_letter(column)].hidden = True
''', '''    write_title(sheet, f"{CF_TITLE} - {code} {year}")
    for column in (CF_LINE_COLUMN, CF_VAT_COLUMN):
        sheet.column_dimensions[get_column_letter(column)].hidden = True

    # This country's rate, looked up rather than written in, so the sheet says where its own
    # number came from. It comes back nought for a country with no VAT, which is what empties
    # both VAT lines on the four American sheets without them needing a different layout.
    vat_rate = f'SUMIFS({VAT_COUNTRY_TABLE}[Rate],{VAT_COUNTRY_TABLE}[Country],"{code}")'
    # The basis this sheet is on, where the note would go: =Basis, so the dropdown on
    # Dashboard CF drives it, and a reader sees which basis the figures are on. Type a basis
    # over it to pin this one sheet. The two hidden cells beside it are read by every column.
    sheet.cell(row=NOTE_ROW, column=CF_LABEL_COLUMN, value=f"={BASIS_NAME}").font = NOTE
    sheet.cell(row=NOTE_ROW, column=CF_LINE_COLUMN,
               value=f'=IF({CF_BASIS_CELL}="{BASIS_DIFFERENCE}",1,0)')
    sheet.cell(row=NOTE_ROW, column=CF_VAT_COLUMN, value=f"={vat_rate}")
''')

replace(CF, '''    sheet.row_dimensions[CF_KEY_ROW].hidden = True
    sheet.sheet_properties.outlinePr.summaryRight = True
''', '''    sheet.row_dimensions[CF_KEY_ROW].hidden = True
    write_cf_basis_rows(sheet, code, year, months, weeks)
    sheet.sheet_properties.outlinePr.summaryRight = True
''')

replace(CF, '''    # This country's rate, looked up rather than written in, so the sheet says where its own
    # number came from. It comes back nought for a country with no VAT, which is what empties
    # both VAT lines on the four American sheets without them needing a different layout.
    vat_rate = f'SUMIFS({VAT_COUNTRY_TABLE}[Rate],{VAT_COUNTRY_TABLE}[Country],"{code}")'
    for key, text, kind, reads in CASH_ROWS:
''', '''    plain_key_row = CF_BASIS_ROW_AT["plain_key"]
    for key, text, kind, reads in CASH_ROWS:
''')

# The week-column formulas: the untouched rows read the plain key, the three replaced rows
# get their own formula, and the "if" variant wraps everything instead.
replace(CF, '''        for month, _ in months:
            for position, _ in weeks[month]:
                letter = get_column_letter(position)
                if kind in ("traded", "memo"):
                    formula = (f'=SUMIFS({PAYMENT_TABLE}[Cash amount],'
                               f'{PAYMENT_TABLE}[Cash key],"{code}|"&$'
                               f'{get_column_letter(CF_LINE_COLUMN)}{row}&"|"&'
                               f'{letter}${CF_KEY_ROW})')
                elif kind == "cost":
                    formula = (f'=SUMIFS({COST_CASH_TABLE}[Cash amount],'
                               f'{COST_CASH_TABLE}[Cash key],"{code}|"&$'
                               f'{get_column_letter(CF_LINE_COLUMN)}{row}&"|"&'
                               f'{letter}${CF_KEY_ROW})')
                elif kind == "vat_in":
''', '''        for month, _ in months:
            for position, _ in weeks[month]:
                letter = get_column_letter(position)
                # The week the untouched rows look up. Under "key" it is the helper row,
                # which is the real week or a key that finds nothing; under "if" it is the
                # real week and the IF below does the switching.
                key_ref = (f"{letter}${plain_key_row}" if CF_SWITCH == "key"
                           else f"{letter}${CF_KEY_ROW}")
                if key in CF_OVERRIDDEN:
                    formula = cf_override_formula(key, letter)
                elif kind in ("traded", "memo"):
                    formula = (f'=SUMIFS({PAYMENT_TABLE}[Cash amount],'
                               f'{PAYMENT_TABLE}[Cash key],"{code}|"&$'
                               f'{get_column_letter(CF_LINE_COLUMN)}{row}&"|"&'
                               f'{key_ref})')
                elif kind == "cost":
                    formula = (f'=SUMIFS({COST_CASH_TABLE}[Cash amount],'
                               f'{COST_CASH_TABLE}[Cash key],"{code}|"&$'
                               f'{get_column_letter(CF_LINE_COLUMN)}{row}&"|"&'
                               f'{key_ref})')
                elif kind == "vat_in":
''')

replace(CF, '''                             f'*${share_letter}${CASH_ROW_AT[member]}'
                             for member in reads]
                    formula = f"=({'+'.join(terms)})*{vat_rate}"
''', '''                             f'*${share_letter}${CASH_ROW_AT[member]}'
                             for member in reads]
                    formula = f"=({'+'.join(terms)})*{CF_RATE_CELL}"
''')
replace(CF, '''                    formula = (f"=SUMPRODUCT(${share_letter}${top}:${share_letter}${bottom},"
                               f"{letter}${top}:{letter}${bottom})*{vat_rate}")
''', '''                    formula = (f"=SUMPRODUCT(${share_letter}${top}:${share_letter}${bottom},"
                               f"{letter}${top}:{letter}${bottom})*{CF_RATE_CELL}")
''')
replace(CF, '''                               f'"{code}|{VAT_SETTLEMENT_LINE}|"&{letter}${CF_KEY_ROW})')
''', '''                               f'"{code}|{VAT_SETTLEMENT_LINE}|"&{key_ref})')
''')
replace(CF, '''                else:
                    formula = "=" + "+".join(f"{letter}{CASH_ROW_AT[member]}"
                                             for member in reads)
                write_cf_cell(sheet, row, position, kind, formula)
''', '''                else:
                    formula = "=" + "+".join(f"{letter}{CASH_ROW_AT[member]}"
                                             for member in reads)
                if (CF_SWITCH == "if" and key not in CF_OVERRIDDEN
                        and kind in ("traded", "memo", "cost", "vat_in", "settlement")):
                    formula = (f'=IF({letter}${CF_BASIS_ROW_AT["mult"]}=1,'
                               f'{formula[1:]},0)')
                write_cf_cell(sheet, row, position, kind, formula)
''')

# --- Dashboard CF: the switch itself -------------------------------------------------------
replace(DASH, '''DASH_CF_SHEET = "Dashboard CF"
DASH_CF_TITLE = "Cash flow, month by month"
''', '''DASH_CF_SHEET = "Dashboard CF"
DASH_CF_TITLE = "Cash flow, month by month"
DASH_SWITCH_ROW = DASH_COMPARE_ROW + 1     # the basis dropdown, under the two selections
''')
replace(DASH, '''    add_dropdown(sheet, "E", CASH_YEARS, DASH_SELECT_ROW, DASH_COMPARE_ROW)

    # --- what the block is showing
    caption = sheet.cell(row=DASH_CAPTION_ROW, column=DASH_LABEL_COLUMN,
                         value=f'={COUNTRY_CELL}&" "&{YEAR_CELL}&"  -  {DASH_CF_TITLE}"')
''', '''    add_dropdown(sheet, "E", CASH_YEARS, DASH_SELECT_ROW, DASH_COMPARE_ROW)

    # The basis every CF sheet is on, and the one cell the overlay is switched from. Named,
    # so the twelve sheets read it by name rather than by address. Human-owned: it is on the
    # rebuild checklist in the README.
    sheet.cell(row=DASH_SWITCH_ROW, column=2, value="Basis").font = HEADER
    cell = sheet.cell(row=DASH_SWITCH_ROW, column=3, value=BASIS_BUDGET)
    cell.font, cell.protection = ENTRY, Protection(locked=False)
    add_dropdown(sheet, "C", BASES, DASH_SWITCH_ROW, DASH_SWITCH_ROW)
    book.defined_names[BASIS_NAME] = DefinedName(
        BASIS_NAME, attr_text=f"'{DASH_CF_SHEET}'!$C${DASH_SWITCH_ROW}")

    # --- what the block is showing
    caption = sheet.cell(row=DASH_CAPTION_ROW, column=DASH_LABEL_COLUMN,
                         value=f'={COUNTRY_CELL}&" "&{YEAR_CELL}&"  -  {DASH_CF_TITLE}'
                               f'  -  "&{BASIS_NAME}')
''')

# --- Short-term input: name the first week of the window ----------------------------------
replace(ST, '''    anchor_cell, keys = write_st_window(sheet, window_first, st_anchor_monday())
''', '''    anchor_cell, keys = write_st_window(sheet, window_first, st_anchor_monday())
    # The boundary the cash flow sheets draw: every week before this one is the past under
    # an override, and shows nothing (ticket 01). Named, so the sheets read it by name.
    book.defined_names[FIRST_PASTE_NAME] = DefinedName(
        FIRST_PASTE_NAME, attr_text=f"'{ST_SHEET}'!{keys[0]}")
''')
replace(ST, '''# CFO reads a rectangle and does not read a twenty-row table. Nothing reads this sheet yet;
# ticket 05 wires it into the cash flow.
''', '''# CFO reads a rectangle and does not read a twenty-row table. The cash flow sheets read the
# `ShortTerm` table on the right of it, under the basis switch on Dashboard CF (ticket 05).
''')

json.dump(nb, io.open(NB, "w", encoding="utf-8"), indent=1, ensure_ascii=False)
io.open(NB, "a", encoding="utf-8").write("\n")
print("patched cells", IMPORTS, CF, DASH, ST)
