"""Ticket 09: grey the weeks with no figure on this basis, in forecast_dev.ipynb.

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

if "CF_UNSHOWN" in src(CF):
    print("already patched")
    sys.exit(0)

# --- imports ---------------------------------------------------------------------------
replace(IMPORTS,
        "from openpyxl.chart.text import RichText\n",
        "from openpyxl.chart.text import RichText\n"
        "from openpyxl.formatting.rule import Rule\n")
replace(IMPORTS,
        "from openpyxl.styles import Alignment, Border, Font, Protection, Side\n",
        "from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side\n"
        "from openpyxl.styles.differential import DifferentialStyle\n"
        "from openpyxl.styles.numbers import NumberFormat\n")

# --- the helper row ----------------------------------------------------------------------
replace(CF, """    ("mult",        "Budget figure survives"),      # live, and not Difference
""", """    ("mult",        "Budget figure survives"),      # live, and not Difference
    ("shown",       "Figure on this basis"),        # mult or override; the greying keys off it
""")

replace(CF, """                "mult": f'={ref("live")}*(1-{CF_DIFF_CELL})',
""", """                "mult": f'={ref("live")}*(1-{CF_DIFF_CELL})',
                "shown": f'=IF(OR({ref("mult")}=1,{ref("override")}=1),1,0)',
""")

# --- the style, beside CF_MONEY ----------------------------------------------------------
replace(CF, """CF_MONEY = '#,##0,;(#,##0,);"-"'
""", """CF_MONEY = '#,##0,;(#,##0,);"-"'

# How a week with no figure on this basis reads: a past week under an override, and every
# week but the pasted ones under Difference. Greyed with its header, and empty rather than
# `-`, so it cannot be mistaken for a week with no cash; the note cell says why, once. A
# format keyed off the `shown` row rather than a change to the figures, so the cells still
# hold nought and the `+`/`-` totals keep working. A month with no shown week is greyed with
# its weeks; a month partly shown totals the weeks it shows, which is what SUM does. Built as
# lettered variants and compared in Excel (ticket 09).
CF_UNSHOWN_FILL = PatternFill(fill_type="solid", start_color=rgb(GRID), end_color=rgb(GRID))
CF_UNSHOWN = DifferentialStyle(fill=CF_UNSHOWN_FILL,
                               numFmt=NumberFormat(numFmtId=190, formatCode=";;;"))
CF_UNSHOWN_HEAD = DifferentialStyle(fill=CF_UNSHOWN_FILL, font=Font(color=rgb(MUTED)))
""")

# --- the note cell says which weeks are not shown -------------------------------------------
replace(CF, """    sheet.cell(row=NOTE_ROW, column=CF_LABEL_COLUMN,
               value=f'="Basis: "&{BASIS_NAME}').font = NOTE
""", """    sheet.cell(row=NOTE_ROW, column=CF_LABEL_COLUMN,
               value=f'="Basis: "&{BASIS_NAME}'
                     f'&IF({BASIS_NAME}="{BASIS_BUDGET}","",'
                     f'IF({BASIS_NAME}="{BASIS_DIFFERENCE}",'
                     f'"   |   only the pasted weeks are shown",'
                     f'"   |   weeks before "&{FIRST_PASTE_NAME}&" are not shown"))').font = NOTE
""")

# --- the greying, after the lines are written ---------------------------------------------
replace(CF, """    sheet.freeze_panes = f"{get_column_letter(CF_FIRST_DATA_COLUMN)}{CF_FIRST_ROW}"
    return sheet
""", """    grey_unshown_weeks(sheet, months, weeks, totals)
    sheet.freeze_panes = f"{get_column_letter(CF_FIRST_DATA_COLUMN)}{CF_FIRST_ROW}"
    return sheet


def grey_unshown_weeks(sheet, months, weeks, totals):
    \"\"\"Grey every column with no figure on this basis, header and body, month by month.

    One rule per month's run of weeks, off the first week's `shown` cell and relative from
    there, and one per month total, off the whole run; the year column is never greyed.
    \"\"\"
    shown_row = CF_BASIS_ROW_AT["shown"]
    last_row = max(CASH_ROW_AT.values())
    for month, _ in months:
        first = get_column_letter(weeks[month][0][0])
        last = get_column_letter(weeks[month][-1][0])
        total = get_column_letter(totals[month])
        run, whole = f"{first}${shown_row}=0", f"SUM({first}${shown_row}:{last}${shown_row})=0"
        for columns, formula in ((f"{first}:{last}", run), (f"{total}:{total}", whole)):
            left, right = columns.split(":")
            sheet.conditional_formatting.add(
                f"{left}{CF_HEADER_ROW}:{right}{CF_HEADER_ROW}",
                Rule(type="expression", formula=[formula], dxf=CF_UNSHOWN_HEAD))
            sheet.conditional_formatting.add(
                f"{left}{CF_FIRST_ROW}:{right}{last_row}",
                Rule(type="expression", formula=[formula], dxf=CF_UNSHOWN))
""")

json.dump(nb, io.open(NB, "w", encoding="utf-8"), indent=1, ensure_ascii=False)
io.open(NB, "a", encoding="utf-8").write("\n")
print("patched")
