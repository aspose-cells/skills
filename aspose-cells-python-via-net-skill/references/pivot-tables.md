# Pivot Tables

Read this when creating a pivot table, changing its source range or fields, or when a
pivot shows no data, only headers, or stale (previous-run) numbers.

Aspose.Cells writes the pivot *definition* plus a cached snapshot of values; it never
recomputes the body automatically. Every value in the pivot area is stale until you
calculate. `calculate_data()` computes the view from the inner *pivot cache*; it does not
re-read the source range. If the source data changed, refresh the cache first.

## Pitfalls

### Pivot body is empty, shows only headers, or shows stale numbers

Symptom: after `pivot_tables.add(...)`, or after editing the source cells, the pivot area is
blank, shows field headers with no values, or keeps the numbers from a previous run.
Cause: the body is computed on demand from the pivot cache. `add` builds the cache once
from the source at that moment; later edits to the source cells do not flow into the cache,
and nothing recomputes on save.
Fix, in order:
1. If the source cells are formula-driven, call `workbook.calculate_formula()` first so the
   source values are current (Aspose does not auto-evaluate formulas either).
2. Rebuild the cache from the source and recompute the view in one non-obsolete call:
   `opt = PivotTableCalculateOption(); opt.refresh_data = True; pt.calculate_data(opt)`.
3. Do this before `save()` and before reading any pivot output cell.

```python
import aspose.cells as gc
import aspose.cells.pivot as pv

workbook = gc.Workbook("report.xlsx")
pt = workbook.worksheets[0].pivot_tables[0]
workbook.calculate_formula()          # only needed if the source cells are formulas
opt = pv.PivotTableCalculateOption()
opt.refresh_data = True
pt.calculate_data(opt)                # rebuild cache from current source, then recompute
workbook.save("report.xlsx")
```

`pt.calculate_data()` with no option recomputes only from the *existing* cache. That is fine
immediately after `add` (the cache was just built), but it will NOT pick up later edits to
the source cells.

To refresh the cache from the current source, then recompute the view, use the one-call
form above. The two-step equivalent `pt.pivot_cache.refresh(); pt.calculate_data()` also works -
`pivot_cache` is exposed as `PivotTable.pivot_cache`. `calculate_data()` itself is never obsolete -
only the *refresh* step changed: the old `pt.refresh_data()` is obsolete in 26.7 (marked obsolete -
use `pt.pivot_cache.refresh()` instead).

### Cannot compute correctly in code - defer refresh to Excel

Symptom: the pivot depends on a source you cannot fully reproduce headless, or you want the
numbers to be live when the user opens the file.
Cause: you do not want to bake a computed snapshot at save time.
Fix: set `pt.refresh_data_on_opening_file = True`. Excel then refreshes the pivot when the file
is opened. Note this does not populate the body inside the saved file, so a headless reader
(including Aspose re-opening it) still sees the old cache until it calculates.

### Reading pivot output cells returns null before calculation

Symptom: reading cells in the pivot area returns null/empty even though the pivot looks set up.
Cause: pivot output cells are ordinary cells only *after* the body is computed.
Fix: call `calculate_data` (see above), then read via the pivot's ranges. `data_body_range` is
the values region; `table_range1` is the whole report excluding page fields; `table_range2`
includes page fields. All are `CellArea` with `start_row`/`start_column`/`end_row`/`end_column`.

```python
import aspose.cells as gc

workbook = gc.Workbook("report.xlsx")
sheet = workbook.worksheets[0]
pt = sheet.pivot_tables[0]
pt.calculate_data()                                  # body cells are null before this
body = pt.data_body_range                           # values region
v = sheet.cells.get(body.start_row, body.start_column).value
```

### Recomputing stale output after source edits

Symptom: the pivot still shows old numbers after you change the source cells.
Cause: the pivot body is a snapshot computed from the cache; editing the source does not
recompute it.
Fix: call `calculate_data(opt)` (or the two-step `pt.pivot_cache.refresh(); calculate_data()`) after
source edits and before reading/saving. `pivot_cache` is public (`pt.pivot_cache.refresh()`), and
`get_dependent_pivot_tables()` exists for child pivots.

## Create a pivot table

`PivotTables.add(source, dest_cell, name)` returns the new pivot's index. `source` is a
range string like `"=Sheet1!A1:C100"` (top-left to bottom-right; reversed ranges are
invalid) or a named range. Add fields by their source column index into an area with
`add_field_to_area(PivotFieldType, column_index)`.

```python
import aspose.cells as gc
import aspose.cells.pivot as pv

sheet = gc.Workbook().worksheets[0]
# source data assumed to live in Sheet1!A1:C100 with headers in row 1
idx = sheet.pivot_tables.add("=Sheet1!A1:C100", "F3", "SalesPivot")
pt = sheet.pivot_tables[idx]
pt.add_field_to_area(pv.PivotFieldType.ROW, 0)      # 1st source column -> Row area
pt.add_field_to_area(pv.PivotFieldType.COLUMN, 1)   # 2nd column -> Column area
pt.add_field_to_area(pv.PivotFieldType.DATA, 2)     # 3rd column -> Data (values) area
pt.calculate_data()                                 # compute the body before save/read
```

`PivotFieldType` values: `ROW`, `COLUMN`, `DATA`, `PAGE`, `UNDEFINED`. `add_field_to_area`
also has overloads taking a field name (`str`) or a `PivotField`. There are `add`
overloads with an `int row, int column` destination instead of a cell string.

## Aggregation function and number format on a data field

`data_fields` is a collection; each entry is a `PivotField`. Set its `function`
(`ConsolidationFunction`) and `number_format` (an Excel format string).

```python
import aspose.cells as gc

pt = gc.Workbook("data.xlsx").worksheets[0].pivot_tables[0]
data0 = pt.data_fields[0]
data0.function = gc.ConsolidationFunction.SUM    # AVERAGE, COUNT, MAX, MIN, PRODUCT, ...
data0.number_format = "#,##0.00"
data0.display_name = "Total Sales"
```

`ConsolidationFunction` includes `SUM`, `AVERAGE`, `COUNT`, `COUNT_NUMS`, `DISTINCT_COUNT`,
`MAX`, `MIN`, `PRODUCT`, `STD_DEV`, `STD_DEVP`, `VAR`, `VARP` (`DISTINCT_COUNT` needs Excel
2013+).

## Page filter, grand totals, style, empty-cell text

```python
import aspose.cells as gc
import aspose.cells.pivot as pv

pt = gc.Workbook("data.xlsx").worksheets[0].pivot_tables[0]
pt.add_field_to_area(pv.PivotFieldType.PAGE, 3)     # report (page) filter field
pt.show_row_grand_totals = True
pt.show_column_grand_totals = True
pt.pivot_table_style_type = pv.PivotTableStyleType.PIVOT_TABLE_STYLE_MEDIUM9
pt.display_null_string = True
pt.null_string = "-"                           # "for empty cells show" text
```

Layout form: `pt.show_in_compact_form()`, `pt.show_in_outline_form()`, `pt.show_in_tabular_form()`.
Style names run `PIVOT_TABLE_STYLE_LIGHT1..28`, `...MEDIUM1..28`, `...DARK1..28`, plus `NONE`.

## Manual cell formatting

`format_all(style)` styles the whole pivot; `format(row, col, style)` styles one cell. Build
the `Style` from the workbook, never with a bare constructor.

```python
import aspose.cells as gc
import aspose.pydrawing as drawing

workbook = gc.Workbook("data.xlsx")
pt = workbook.worksheets[0].pivot_tables[0]
style = workbook.create_style()
style.foreground_color = drawing.Color.light_yellow
style.pattern = gc.BackgroundType.SOLID
pt.format_all(style)                           # format(r, c, style) for a single cell
```

## Calculated field

`add_calculated_field(name, formula, drag_to_data_area)` adds a field computed from other fields;
pass `True` to drop it into the Data area immediately. Recompute afterward.

```python
import aspose.cells as gc
import aspose.cells.pivot as pv

pt = gc.Workbook("data.xlsx").worksheets[0].pivot_tables[0]
pt.add_calculated_field("Margin", "=Revenue-Cost", True)   # True -> into Data area
opt = pv.PivotTableCalculateOption()
opt.refresh_data = True
pt.calculate_data(opt)
```

## Change the data source

`change_data_source(string[])` repoints the pivot; pass a one-element array for a single
range. Recompute after changing it.

```python
import aspose.cells as gc
import aspose.cells.pivot as pv

pt = gc.Workbook("data.xlsx").worksheets[0].pivot_tables[0]
pt.change_data_source(["=Data!A1:D500"])    # top-left to bottom-right
opt = pv.PivotTableCalculateOption()
opt.refresh_data = True
pt.calculate_data(opt)
```

## Refresh nested / child pivots

A pivot whose source is another pivot is a child. Refresh the parent first, then each child
returned by `get_dependent_pivot_tables()` (the old `get_children()` is obsolete in 26.7).

```python
import aspose.cells as gc
import aspose.cells.pivot as pv

workbook = gc.Workbook("nested.xlsx")
parent = workbook.worksheets[0].pivot_tables[0]
opt = pv.PivotTableCalculateOption()
opt.refresh_data = True
parent.calculate_data(opt)                                # parent cache first
for child in parent.get_dependent_pivot_tables():
    child.calculate_data(opt)                             # then each child
```

## Sort, filter, delete

Auto-sort a field on its values: `f = pt.row_fields[0]; f.is_auto_sort = True; f.is_ascend_sort = True`.
Clear all field filters with `pt.clear_filters()`. Delete a pivot with
`sheet.pivot_tables.remove_at(index)` or `sheet.pivot_tables.remove(pt)`.

## Pivot charts

Bind a chart to a pivot via the chart's pivot source, then refresh the pivot and the chart
before rendering. See charts.md.

## Related
- formulas-and-calculation.md
- charts.md
- cell-values-and-data.md
- styles-and-formatting.md
- worksheets-rows-columns.md
