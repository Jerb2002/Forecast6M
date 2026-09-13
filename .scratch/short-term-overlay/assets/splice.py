"""Splice the short-term input section into forecast_dev.ipynb, just before the build cell."""
import io, json

NB = "forecast_dev.ipynb"
MARKS = ("PROTOTYPE - ticket 02",                    # the first pass, both variants
         "The short-term paste - ticket 02",         # this pass
         "### The short-term input sheet")

md = io.open(".scratch/short-term-overlay/assets/st_markdown.md", encoding="utf-8").read()
code = io.open(".scratch/short-term-overlay/assets/st_code.py", encoding="utf-8").read()

nb = json.load(io.open(NB, encoding="utf-8"))

# Idempotent: drop any copy of this section left by an earlier run.
cells = [c for c in nb["cells"]
         if not any(mark in "".join(c["source"]) for mark in MARKS)]

# The build cell is the one that saves the workbook.
build = next(i for i, c in enumerate(cells)
             if c["cell_type"] == "code" and "book.save(WORKBOOK)" in "".join(c["source"]))


def lines(text):
    """nbformat wants a list of lines, each keeping its newline except the last."""
    parts = text.rstrip("\n").split("\n")
    return [p + "\n" for p in parts[:-1]] + [parts[-1]]


cells[build:build] = [
    {"cell_type": "markdown", "metadata": {}, "source": lines(md)},
    {"cell_type": "code", "execution_count": None, "metadata": {}, "outputs": [],
     "source": lines(code)},
]

# Wire the writer into the build, at the end of the Inputs group. The earlier pass wrote two
# calls here and one of those sheets no longer exists, so the whole block is replaced.
build += 2
src = "".join(cells[build]["source"])
anchor = "write_cost_timing(book)"
old = [line for line in src.split("\n")
       if "write_short_term" in line or "Prototype, ticket 02" in line]
for line in old:
    src = src.replace(line + "\n", "")
assert anchor in src, "could not find write_cost_timing(book) in the build cell"
src = src.replace(anchor, anchor + "\nwrite_short_term(book)", 1)
cells[build]["source"] = lines(src)

nb["cells"] = cells
json.dump(nb, io.open(NB, "w", encoding="utf-8", newline="\n"), indent=1, ensure_ascii=False)
print("spliced; removed", len(old), "old build lines; build cell is now", build)
