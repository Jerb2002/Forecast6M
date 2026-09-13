### The short-term input sheet

Every Monday the short-term tool produces five weeks of revenue cash per country. Those
figures have to reach this workbook without the notebook being run again, which means they
land on a sheet somebody pastes into. What that sheet looks like was settled by building two
of them and opening both in Excel, rather than by arguing about it.

**Wide won.** A rectangle - a row per country, five week columns across - is what a CFO
reads, and a twenty-row table is not, however much the rest of this workbook is built that
way. The cost is real and it is on the right of the sheet: every figure on a cash flow sheet
is looked up by one key per row, `DNK|Revenue|2026-09-W37`, so the rectangle is unpivoted into
twenty rows before anything can read it. Those twenty rows are formulas back into the paste,
so they cannot drift from it, but they are a second block on a sheet that would otherwise
have one.

**The window** is one cell: the Monday the five weeks start on. Everything else is worked out
from it, including the week each column is headed with - so the five weeks cannot be
half-updated. The columns label themselves, which they could not do inside a real Excel table,
because a table will not hold a formula in its header. The paste block is therefore a styled
range and the unpivot beside it is the named table.

**What the shape cannot do** is label the amounts. A figure in a rectangle knows only which
column it sits in, so moving the Monday without re-pasting slides last week's figures onto
this week - quietly, and with the sheet reporting the window as current. The long shape caught
exactly this, and losing that catch is the price of the shape that gets used. The guard that
remains is the window block saying how many weeks old the Monday is, which catches the other
failure: a Monday nobody touched at all. The order of the two steps is therefore part of the
procedure rather than a detail, and it is written on the sheet.

A pasted week is dated by its **Thursday**, which is what decides both the ISO week and the
year that week belongs to. That used to be a simplification this sheet made on its own, because
the cash flow keyed the month and the week together and a straddling week had a column under
each month. It no longer does: the cash flow is keyed at the week level, so this sheet and every
other one name a week the same way.
