"""Run forecast_dev.ipynb's code cells in order, then check the week columns hold together."""
import io, json, sys
import matplotlib
matplotlib.use("Agg")          # no windows; the charts are not what this run is for

nb = json.load(io.open("forecast_dev.ipynb", encoding="utf-8"))
env = {"__name__": "__main__"}
for index, cell in enumerate(nb["cells"]):
    if cell["cell_type"] != "code":
        continue
    src = "".join(cell["source"])
    if src.lstrip().startswith(("%", "!")):
        continue
    try:
        exec(compile(src, f"<cell {index}>", "exec"), env)
    except Exception:
        import traceback
        print(f"\n*** cell {index} failed ***", file=sys.stderr)
        traceback.print_exc()
        sys.exit(1)

print("OK - sheets:", len(env["book"].sheetnames))

# --- the two things the re-keying could break -------------------------------------------
cash_year_weeks, years = env["cash_year_weeks"], env["CASH_YEARS"]
cash_week_of = env["cash_week_of"]

seen, duplicates = {}, []
for year in years:
    for month, weeks in cash_year_weeks(year):
        for key in weeks:
            if key in seen:
                duplicates.append((key, seen[key], (year, month)))
            seen[key] = (year, month)

print("week columns:", len(seen), "over", len(years), "years")
print("duplicate week columns:", duplicates if duplicates else "none")

months_per_year = {year: len(cash_year_weeks(year)) for year in years}
print("month groups per year:", months_per_year,
      "(all 12)" if set(months_per_year.values()) == {12} else "*** NOT ALL 12 ***")

missing = set()
for frame, column in ((env["payments"], "payment_date"), (env["costs"], "paid_date")):
    for date in frame[column]:
        key = cash_week_of(date)
        if key not in seen:
            missing.add(key)
print("cash weeks with no column:", sorted(missing) if missing else "none")

if duplicates or missing or set(months_per_year.values()) != {12}:
    sys.exit(1)
