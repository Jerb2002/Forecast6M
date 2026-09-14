"""Ticket 10: pin `Check` to the budget basis and add its VAT block, in forecast_dev.ipynb.

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


def find(mark, kind="code"):
    hits = [i for i, c in enumerate(cells) if c["cell_type"] == kind and mark in src(i)]
    assert len(hits) == 1, (mark, hits)
    return hits[0]


def replace(i, old, new, count=1):
    s = src(i)
    assert s.count(old) == count, (i, old[:60], s.count(old))
    put(i, s.replace(old, new))


CF = find('CF_SHEET_PREFIX = "CF"')
CHECK = find('CHECK_SHEET = "Check"')
CF_DOC = find("The `Check` sheet ties CF to Budget on the budget basis, so it reads *OK*",
              "markdown")
CHECK_DOC = find("It is a **bridge** rather than two figures and a difference.", "markdown")

if "CF_PURCHASE_KEYS" in src(CF):
    print("already patched")
    sys.exit(0)

# --- CF cell: the hidden rows move below cash_rows(), which they now read --------------
s = src(CF)
start = s.index("# The switch reaches the figures through hidden rows under the key row")
end = s.index('CF_OVERRIDDEN = {"in_revenue": "rev", "in_duty": "duty", "in_vat": "vat"}\n')
end += len('CF_OVERRIDDEN = {"in_revenue": "rev", "in_duty": "duty", "in_vat": "vat"}\n')
block = s[start:end]
s = s[:start] + s[end:]
anchor = "CASH_ROWS = cash_rows()\n"
assert s.count(anchor) == 1
s = s.replace(anchor, anchor + "\n" + block.rstrip("\n") + "\n")
put(CF, s)

replace(CF, """    ("vat_budget",  "VAT on sales, budget"),
    ("vat_direct",  "VAT on sales, direct leg only"),
]
""", """    ("vat_budget",  "VAT on sales, budget"),
    ("vat_direct",  "VAT on sales, direct leg only"),
    # What `Check` reads: the budget-basis figure for every line it ties, keyed on the real
    # week whatever the visible rows are showing (ticket 07). The cash-in side is the three
    # `*_budget` rows above. The cash-out side is a copy of every purchase line, because
    # `VAT on purchases` is each line times its own share and a lumped figure cannot be
    # split back; the cost figure and the VAT are then read off that block.
    *[(f"{key}_budget", f"{text}, budget") for key, text in CF_PURCHASES.items()],
    ("cost_budget",     "Cost cash, budget"),         # the purchase block, added up
    ("vat_out_budget",  "VAT on purchases, budget"),  # the purchase block times its shares
    ("settle_budget",   "VAT settlement, budget"),
    ("fee_budget",      "Factoring fee, budget"),
]
# The hidden rows that carry a month total and a year total, which is what `Check` reads.
CF_BASIS_TOTALLED = ("rev_budget", "duty_budget", "vat_budget", "cost_budget",
                     "vat_out_budget", "settle_budget", "fee_budget")
""")

replace(CF, """CF_BASIS_ROWS = [
    ("live",        "Live on this basis"),          # 1 shown; 0 before the paste window
""", """# The purchase lines: every payment to a supplier, which is what `VAT on purchases` is
# worked out over. Taken from the list rather than named, so a cost line added in section
# 11 gets its hidden copy without anybody remembering to add one.
CF_PURCHASES = {key: text for key, text, kind, _ in CASH_ROWS
                if key and key.startswith("out_") and kind in ("traded", "cost")}
CF_PURCHASE_KEYS = list(CF_PURCHASES)

CF_BASIS_ROWS = [
    ("live",        "Live on this basis"),          # 1 shown; 0 before the paste window
""")

# The formulas: one copy per purchase line on the real week, and the four rows off them.
replace(CF, """def write_cf_basis_rows(sheet, code, year, months, weeks):
    \"\"\"The hidden rows the basis reaches the figures through, one figure per week column.

    `ShortTerm` and the budget table are both small, so the lookups here are cheap; what
    costs is the six budget-basis rows, which are the same lookups the three overridden
    lines used to carry themselves, with a `Direct`-leg copy of each beside it. That is the
    price of ticket 03: an overridden week keeps the budget's factored legs and drops only
    the leg the customer pays, so the sheet has to be able to tell the two apart.
    \"\"\"
    at = CF_BASIS_ROW_AT
    line_letter = get_column_letter(CF_LINE_COLUMN)
    share_letter = get_column_letter(CF_VAT_COLUMN)
    rev_row, duty_row = CASH_ROW_AT["in_revenue"], CASH_ROW_AT["in_duty"]
    v = f"${share_letter}${rev_row}*{CF_RATE_CELL}"      # the VAT rate on the service part
    direct = f'{PAYMENT_TABLE}[Leg],"{LEG_LABELS["direct"]}"'

    for key, text in CF_BASIS_ROWS:
        sheet.row_dimensions[at[key]].hidden = True
        sheet.cell(row=at[key], column=CF_LABEL_COLUMN, value=text).font = NOTE
""", """def write_cf_basis_rows(sheet, code, year, months, weeks, totals, year_column):
    \"\"\"The hidden rows the basis reaches the figures through, one figure per week column.

    `ShortTerm` and the budget table are both small, so the lookups here are cheap; what
    costs is the budget-basis rows, which are the same lookups the visible lines carry,
    keyed on the real week rather than the plain key. Six of them are the price of ticket
    03: an overridden week keeps the budget's factored legs and drops only the leg the
    customer pays, so the sheet has to be able to tell the two apart. The rest are the
    price of ticket 07: `Check` ties the budget basis whatever the dropdown says, and the
    visible cost rows are nought in a week that is not shown, so it reads copies instead.
    Those copies carry month and year totals, the way the visible rows do, so `Check` reads
    the year column of a hidden row exactly as it read the year column of a visible one.
    \"\"\"
    at = CF_BASIS_ROW_AT
    line_letter = get_column_letter(CF_LINE_COLUMN)
    share_letter = get_column_letter(CF_VAT_COLUMN)
    rev_row, duty_row = CASH_ROW_AT["in_revenue"], CASH_ROW_AT["in_duty"]
    v = f"${share_letter}${rev_row}*{CF_RATE_CELL}"      # the VAT rate on the service part
    direct = f'{PAYMENT_TABLE}[Leg],"{LEG_LABELS["direct"]}"'
    top, bottom = (at[f"{CF_PURCHASE_KEYS[0]}_budget"],
                   at[f"{CF_PURCHASE_KEYS[-1]}_budget"])

    for key, text in CF_BASIS_ROWS:
        sheet.row_dimensions[at[key]].hidden = True
        sheet.cell(row=at[key], column=CF_LABEL_COLUMN, value=text).font = NOTE
    for key in CF_PURCHASE_KEYS:
        # The share of this line that carries VAT, beside its copy as it is beside the
        # line, so the VAT copy below is the same SUMPRODUCT the visible row is.
        sheet.cell(row=at[f"{key}_budget"], column=CF_VAT_COLUMN,
                   value=f"=${share_letter}${CASH_ROW_AT[key]}")
""")

replace(CF, """            def traded(row, leg=""):
                return (f'SUMIFS({PAYMENT_TABLE}[Cash amount],{PAYMENT_TABLE}[Cash key],'
                        f'"{code}|"&${line_letter}${row}&"|"&{week}{leg})')
""", """            def traded(row, leg="", table=PAYMENT_TABLE):
                return (f'SUMIFS({table}[Cash amount],{table}[Cash key],'
                        f'"{code}|"&${line_letter}${row}&"|"&{week}{leg})')
""")

replace(CF, """                "vat_budget": "=" + vat(),
                "vat_direct": "=" + vat("," + direct),
            }
            for key, formula in formulas.items():
                sheet.cell(row=at[key], column=position, value=formula)
""", """                "vat_budget": "=" + vat(),
                "vat_direct": "=" + vat("," + direct),
                "cost_budget": f"=SUM({letter}${top}:{letter}${bottom})",
                "vat_out_budget": (f"=SUMPRODUCT(${share_letter}${top}:${share_letter}${bottom},"
                                   f"{letter}${top}:{letter}${bottom})*{CF_RATE_CELL}"),
                "settle_budget": (f'=SUMIFS({VAT_SETTLEMENT_TABLE}[Net],'
                                  f'{VAT_SETTLEMENT_TABLE}[Cash key],'
                                  f'"{code}|{VAT_SETTLEMENT_LINE}|"&{week})'),
                "fee_budget": "=" + traded(CASH_ROW_AT["out_fee"]),
            }
            for key in CF_PURCHASE_KEYS:
                # The line's own lookup on the real week: COGS and its duty off the payment
                # sheet, the scheduled lines off the cost cash sheet, as the visible row is.
                row = CASH_ROW_AT[key]
                table = PAYMENT_TABLE if key in ("out_cogs", "out_cogs_duty") else COST_CASH_TABLE
                formulas[f"{key}_budget"] = "=" + traded(row, table=table)
            for key, formula in formulas.items():
                sheet.cell(row=at[key], column=position, value=formula)

    # The month, then the year off the months, for the rows `Check` reads.
    month_letters = [get_column_letter(totals[month]) for month, _ in months]
    for key in CF_BASIS_TOTALLED:
        row = at[key]
        for month, _ in months:
            first = get_column_letter(weeks[month][0][0])
            last = get_column_letter(weeks[month][-1][0])
            sheet.cell(row=row, column=totals[month], value=f"=SUM({first}{row}:{last}{row})")
        sheet.cell(row=row, column=year_column,
                   value="=" + "+".join(f"{letter}{row}" for letter in month_letters))
""")

replace(CF, "    write_cf_basis_rows(sheet, code, year, months, weeks)\n",
        "    write_cf_basis_rows(sheet, code, year, months, weeks, totals, year_column)\n")

# --- CF write-up ------------------------------------------------------------------------
replace(CF_DOC, """The `Check` sheet ties CF to Budget on the budget basis, so it reads *OK* under *Budget only*
and *Off* under either of the other two. What it should say then is its own question.
""", """The `Check` sheet ties CF to Budget on the budget basis whatever the dropdown says. It
cannot read the visible rows for that, because under an override they are nought for every
week that is not shown; so each sheet keeps a budget-basis copy of every line the check
ties, in the same hidden block, keyed on the real week. The cost side is one copy per
purchase line rather than one lumped figure, because `VAT on purchases` is each line times
its own share and the check ties the VAT too.
""")

# --- Check cell: read the hidden rows, add the VAT block, say so under the title --------
replace(CHECK, """def cf_year_terms(country, terms):
    \"\"\"Some cash flow rows, added up over every year that country has a sheet for.

    `terms` is (row on the cash flow sheet, sign). The whole of CASH_YEARS is summed and not
    only the budgeted ones, because the tail of a horizon lands in a year nobody budgeted -
    that is the entire reason those later sheets exist, and leaving them out would fail the
    check on money that arrived exactly where it was meant to.
    \"\"\"
    pieces = []
    for year in CASH_YEARS:
        letter, sheet = cf_year_total_column(year), f"'{cf_sheet_name(country, year)}'"
        for key, sign in terms:
            pieces.append(f"{'-' if sign < 0 else '+'}{sheet}!{letter}{CASH_ROW_AT[key]}")
    return "=" + "".join(pieces).lstrip("+")
""", """def cf_year_terms(country, terms):
    \"\"\"Some cash flow rows, added up over every year that country has a sheet for.

    `terms` is (hidden budget-basis row on the cash flow sheet, sign). The hidden rows and
    never the visible ones: the visible figures follow the `Basis` dropdown, and a check
    that read them would say *Off* the moment somebody looked at the override. The hidden
    copies are the same lookups keyed on the real week, so the check still proves the
    sheet's own formulas and not the tables behind them.

    The whole of CASH_YEARS is summed and not only the budgeted ones, because the tail of a
    horizon lands in a year nobody budgeted - that is the entire reason those later sheets
    exist, and leaving them out would fail the check on money that arrived exactly where it
    was meant to.
    \"\"\"
    pieces = []
    for year in CASH_YEARS:
        letter, sheet = cf_year_total_column(year), f"'{cf_sheet_name(country, year)}'"
        for key, sign in terms:
            pieces.append(f"{'-' if sign < 0 else '+'}{sheet}!{letter}{CF_BASIS_ROW_AT[key]}")
    return "=" + "".join(pieces).lstrip("+")
""")

replace(CHECK, """CHECK_OK, CHECK_OFF = "OK", "Off - see the difference rows"
""", """CHECK_OK, CHECK_OFF = "OK", "Off - see the difference rows"
# Said once under the title. The dropdown moves every CF sheet; it does not move this one.
CHECK_NOTE = f"Always on the budget basis, whatever {BASIS_NAME} is set to."
""")

replace(CHECK, """        ("in_fee",     f"Less the {LEG_LABELS['fee'].lower()} the bank keeps", "sheet",
         (("out_fee", -1),)),
""", """        ("in_fee",     f"Less the {LEG_LABELS['fee'].lower()} the bank keeps", "sheet",
         (("fee_budget", -1),)),
""")
replace(CHECK, """        ("in_cash",    "Cash flow sheets, pre-VAT",           "sheet",
         (("in_total", 1), ("in_vat", -1))),
""", """        ("in_cash",    "Cash flow sheets, pre-VAT",           "sheet",
         (("rev_budget", 1), ("duty_budget", 1))),
""")
replace(CHECK, """        ("out_cash",   "Cash flow sheets, pre-VAT",           "sheet",
         (("out_total", 1), ("out_vat", -1), ("out_settle", -1))),
""", """        ("out_cash",   "Cash flow sheets, pre-VAT",           "sheet",
         (("cost_budget", 1),)),
""")
replace(CHECK, """        ("net_diff",   "Difference",                          "diff",
         (("net_expect", 1), ("net_cash", -1))),
        (None,         "",                                    "blank",  None),

        ("verdict",    "Does it tie?",                        "status", ("in_diff", "out_diff")),
""", """        ("net_diff",   "Difference",                          "diff",
         (("net_expect", 1), ("net_cash", -1))),
        (None,         "",                                    "blank",  None),

        # The three VAT rows, over the whole horizon: what we collect, less what we pay, is
        # what we settle, so the net is nought or something is wrong with one of them. The
        # later-year sheets exist so the last settlement lands inside the horizon, which is
        # what lets this tie without leaving a tail out.
        ("vat_head",   "VAT",                                 "title",  None),
        ("vat_in",     "Collected: VAT on sales",             "sheet",
         (("vat_budget", 1),)),
        ("vat_out",    "Paid: VAT on purchases",              "sheet",
         (("vat_out_budget", 1),)),
        ("vat_settle", "Settled: VAT settlement",             "sheet",
         (("settle_budget", 1),)),
        ("vat_net",    "Net",                                 "diff",
         (("vat_in", 1), ("vat_out", -1), ("vat_settle", -1))),
        (None,         "",                                    "blank",  None),

        ("verdict",    "Does it tie?",                        "status",
         ("in_diff", "out_diff", "vat_net")),
""")

replace(CHECK, """            if kind == "status":
                # The same sentence in the total column, where the differences are the four
                # countries added up: a pair of gaps that cancel is still two gaps, and the
                # country columns are where that shows.
                inside = f"{letter}{CHECK_ROW_AT[rule[0]]}"
                outside = f"{letter}{CHECK_ROW_AT[rule[1]]}"
                formula = (f'=IF(AND(ROUND({inside},0)=0,ROUND({outside},0)=0),'
                           f'"{CHECK_OK}","{CHECK_OFF}")')
""", """            if kind == "status":
                # The same sentence in the total column, where the differences are the four
                # countries added up: a pair of gaps that cancel is still two gaps, and the
                # country columns are where that shows.
                tests = ",".join(f"ROUND({letter}{CHECK_ROW_AT[key]},0)=0" for key in rule)
                formula = f'=IF(AND({tests}),"{CHECK_OK}","{CHECK_OFF}")'
""")

replace(CHECK, """    sheet = new_sheet(book, CHECK_SHEET, widths)
    write_title(sheet, CHECK_TITLE)
""", """    sheet = new_sheet(book, CHECK_SHEET, widths)
    write_title(sheet, CHECK_TITLE)
    write_note(sheet, CHECK_NOTE)
""")

# --- Check write-up ---------------------------------------------------------------------
replace(CHECK_DOC, """What is left after those three is the **difference**, and it should be nought. The verdict
row says so in one word per country so nobody has to read four columns of figures to find
out that nothing is wrong.
""", """What is left after those three is the **difference**, and it should be nought.

A fourth block ties the **VAT**, which the bridges above leave out: what we collect on
sales, less what we pay on purchases, is what we settle with the government, so over the
whole horizon the three rows net to nought. The later-year `CF` sheets are what make that
true without a tail left out - the last settlement lands inside them. The block has no
verdict of its own; the one verdict row tests all three differences, and says so in one
word per country so nobody has to read four columns of figures to find out that nothing
is wrong.

**It is always on the budget basis.** The `Basis` dropdown on `Dashboard CF` moves every
`CF` sheet, and a check that read their visible rows would say *Off* the moment somebody
looked at the override - not because anything was wrong, but because a month under a paste
is not meant to add up to the budget. So the sheet reads hidden budget-basis copies each
`CF` sheet keeps of every line it ties, and says so in a note under its title. Those copies
are the sheet's own lookups keyed on the real week, not the tables behind them, so a broken
sheet formula still fails here. How far the paste sits from the budget is not this sheet's
question; *Difference* on the dropdown is.
""")

json.dump(nb, io.open(NB, "w", encoding="utf-8", newline="\n"), indent=1, ensure_ascii=False)
io.open(NB, "a", encoding="utf-8", newline="\n").write("\n")
print("patched")
