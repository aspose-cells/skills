# Slicers and Sparklines

Read this when adding interactive filters (slicers) over a pivot table or Excel table, or when adding/formatting sparklines (the mini line/column charts inside a cell).

Both hang off the worksheet: `worksheet.slicers` (a `SlicerCollection`) and `worksheet.sparkline_groups` (a `SparklineGroupCollection`). Everything else here follows the golden rules - slicers bound to a pivot need the pivot calculated first (see pivot-tables.md), and rendered output needs `calculate()` before `to_image` (see charts.md).

## Slicers

### Attach a slicer to an Excel table (ListObject)

`SlicerCollection.add` returns an int index; index the collection to configure the `Slicer`. The anchor cell name is A1-style (`"E2"`).

```python
import aspose.cells as gc
import aspose.cells.slicers as sl

workbook = gc.Workbook("input.xlsx")
sheet = workbook.worksheets[0]

# Get the existing table (see worksheets-rows-columns.md)
table = sheet.list_objects[0]

# Slicer over the table's column 1 (0-based), anchored at E2
idx = sheet.slicers.add(table, 1, "E2")
slicer = sheet.slicers[idx]

slicer.style_type = sl.SlicerStyleType.SLICER_STYLE_LIGHT1
slicer.caption = "Region"
slicer.number_of_columns = 2   # lay out filter buttons in 2 columns
slicer.column_width = 90       # px per button
workbook.save("out.xlsx", gc.SaveFormat.XLSX)
```

### Attach a slicer to a pivot table

Requires the pivot to exist first (see pivot-tables.md), then add the slicer by the pivot's base field. The slicer works even if the pivot body has not been `calculate_data()`-ed, but refresh the pivot data before rendering anything.

```python
import aspose.cells as gc
import aspose.cells.slicers as sl

workbook = gc.Workbook("report.xlsx")
sheet = workbook.worksheets[0]
pt = sheet.pivot_tables[0]

# Bind to pivot field at 0-based index 0 (or by name: "Region")
idx = sheet.slicers.add(pt, "F2", 0)
slicer = sheet.slicers[idx]
slicer.style_type = sl.SlicerStyleType.SLICER_STYLE_DARK1
```

### Selected items come from the slicer cache

A slicer's allowed values live in `Slicer.slicer_cache` (`SlicerCache.slicer_cache_items`). Each item's `value` is the filter value and `selected` toggles it. The cache is shared between a slicer and its pivot field, so changing it here affects the pivot's filter.

```python
import aspose.cells as gc

workbook = gc.Workbook("report.xlsx")
sheet = workbook.worksheets[0]
slicer = sheet.slicers[0]

items = slicer.slicer_cache.slicer_cache_items
for item in items:
    print(item.value, "selected=", item.selected)
```

### Pitfall - slicer does not filter the sheet until data is refreshed

Symptom: the slicer exists but filtering does nothing to the rows/totals you read.
Cause: a slicer over a pivot only filters the pivot cache; the pivot body (and any formulas over it) must be recalculated after selection changes. Aspose.Cells will not do that for you.
Fix: after changing `selected`, refresh the pivot data and recalculate (see pivot-tables.md), then re-read the affected cells.

```python
import aspose.cells as gc
import aspose.cells.pivot as pv

workbook = gc.Workbook("report.xlsx")
sheet = workbook.worksheets[0]
pt = sheet.pivot_tables[0]
slicer = sheet.slicers[0]

slicer.slicer_cache.slicer_cache_items[0].selected = False

opt = pv.PivotTableCalculateOption()
opt.refresh_data = True
pt.calculate_data(opt)
workbook.calculate_formula()
# ... now read the filtered pivot cells
```

### Pitfall - slicer over a table needs the table, not raw cells

Symptom: `sheet.slicers.add(...)` throws or the slicer shows nothing.
Cause: slicers attach to a `ListObject` or a `PivotTable`, not to an arbitrary cell range. A range that is not a real Excel table has no slicer model.
Fix: create a `ListObject` first (see worksheets-rows-columns.md), or bind to a pivot instead.

## Sparklines

### Add a sparkline group

`SparklineGroups.add(SparklineType, string data_range, bool is_vertical, CellArea location_range)` returns an int index. The data range supplies one row/column of source values per sparkline; the location range is the block of cells that each hold one mini-chart (one cell per series).

```python
import aspose.cells as gc
import aspose.cells.charts as ch

workbook = gc.Workbook()
sheet = workbook.worksheets[0]

for r in range(5):
    sheet.cells.get(r, 0).put_value(r * 2)   # series 1 data, A1:A5
    sheet.cells.get(r, 1).put_value(r * 3)   # series 2 data, B1:B5

# location: D1 holds sparkline for A1:A5, E1 holds sparkline for B1:B5
location = gc.CellArea()
location.start_row = 0
location.start_column = 3
location.end_row = 0
location.end_column = 4
idx = sheet.sparkline_groups.add(ch.SparklineType.LINE, "A1:B5", True, location)
group = sheet.sparkline_groups[idx]
workbook.save("out.xlsx", gc.SaveFormat.XLSX)
```

### Style and per-sparkline access

`SparklineGroup` owns the look for all sparklines in the group; `group.sparklines` indexes the individual sparklines. Preset styles are `SparklinePresetStyleType.STYLE1`..`STYLE36` plus `CUSTOM` - there are no named "Colorful/Dark/Light" members. Colors are set through each point-type's `cells_color` (`high_point_color`, `low_point_color`, `negative_points_color`, `first_point_color`, `last_point_color`, `series_color`); assign `Color` on the `CellsColor`.

```python
import aspose.cells as gc
import aspose.cells.charts as ch
import aspose.pydrawing as drawing

workbook = gc.Workbook()
sheet = workbook.worksheets[0]
for r in range(5):
    sheet.cells.get(r, 0).put_value(r * 2)
    sheet.cells.get(r, 1).put_value(r * 3)

location = gc.CellArea()
location.start_row = 0
location.start_column = 3
location.end_row = 0
location.end_column = 4
idx = sheet.sparkline_groups.add(ch.SparklineType.COLUMN, "A1:B5", True, location)
group = sheet.sparkline_groups[idx]

group.preset_style = ch.SparklinePresetStyleType.STYLE6
group.show_high_point = True
group.high_point_color.color = drawing.Color.red
group.show_low_point = True
group.low_point_color.color = drawing.Color.green
group.show_negative_points = True
group.negative_points_color.color = drawing.Color.orange
group.display_hidden = True

for spark in group.sparklines:
    print(spark.data_range)   # e.g. "Sheet1!$A$1:$A$5"
workbook.save("out.xlsx", gc.SaveFormat.XLSX)
```

Note: `group.sparkline_collection` is obsolete in 26.x - use `group.sparklines`.

### Pitfall - sparkline "shows nothing" in the saved file

Symptom: the sparkline renders as an empty cell, or the mini-chart has no line.
Cause: either the source range is empty, or the location range does not have one cell per data series (Excel requires a 1:1 mapping). A column sparkline with a single value also collapses to a thin line that is easy to miss.
Fix: double-check `CellArea` coordinates (they are 0-based), confirm the data cells actually contain numbers, and make `end_column - start_column + 1` (or rows) equal the number of series.

### Pitfall - a sparkline group for a hidden/blank row

Symptom: sparklines are missing for rows that have no data yet.
Cause: a sparkline is tied to exactly the source cells in its range; blank source cells are simply skipped, they are not "filled down".
Fix: add the sparkline group only after the data is present, or extend the data range to cover the full (possibly empty) range and let the blanks show as gaps.

## Related
- pivot-tables.md - the pivot a slicer filters; refresh before reading filtered output.
- worksheets-rows-columns.md - ListObject tables that slicers attach to.
- charts.md - calculate() before rendering anything.
- styles-and-formatting.md - CellsColor-based colors used by sparkline points.
