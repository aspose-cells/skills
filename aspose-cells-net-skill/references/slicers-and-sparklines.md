# Slicers and Sparklines

Read this when adding interactive filters (slicers) over a pivot table or Excel table, or when adding/formatting sparklines (the mini line/column charts inside a cell).

Both hang off the worksheet: `worksheet.Slicers` (a `SlicerCollection`) and `worksheet.SparklineGroups` (a `SparklineGroupCollection`). Everything else here follows the golden rules - slicers bound to a pivot need the pivot calculated first (see pivot-tables.md), and rendered output needs `Calculate()` before `ToImage` (see charts.md).

## Slicers

### Attach a slicer to an Excel table (ListObject)

`SlicerCollection.Add` returns an int index; index the collection to configure the `Slicer`. The anchor cell name is A1-style (`"E2"`).

```csharp
var workbook = new Workbook("input.xlsx");
Worksheet sheet = workbook.Worksheets[0];

// Get the existing table (see worksheets-rows-columns.md)
ListObject table = sheet.ListObjects[0];

// Slicer over the table's column 1 (0-based), anchored at E2
int idx = sheet.Slicers.Add(table, 1, "E2");
Slicer slicer = sheet.Slicers[idx];

slicer.StyleType = SlicerStyleType.SlicerStyleLight1;
slicer.Caption = "Region";
slicer.NumberOfColumns = 2;   // lay out filter buttons in 2 columns
slicer.ColumnWidth = 90;      // px per button
workbook.Save("out.xlsx", SaveFormat.Xlsx);
```

### Attach a slicer to a pivot table

Requires the pivot to exist first (see pivot-tables.md), then add the slicer by the pivot's base field. The slicer works even if the pivot body has not been `CalculateData()`-ed, but refresh the pivot data before rendering anything.

```csharp
var workbook = new Workbook("report.xlsx");
Worksheet sheet = workbook.Worksheets[0];
PivotTable pt = sheet.PivotTables[0];

// Bind to pivot field at 0-based index 0 (or by name: "Region")
int idx = sheet.Slicers.Add(pt, "F2",0);
Slicer slicer = sheet.Slicers[idx];
slicer.StyleType = SlicerStyleType.SlicerStyleDark1;
```

### Selected items come from the slicer cache

A slicer's allowed values live in `Slicer.SlicerCache` (`SlicerCache.SlicerCacheItems`). Each item's `Value` is the filter value and `Selected` toggles it. The cache is shared between a slicer and its pivot field, so changing it here affects the pivot's filter.

```csharp
var workbook = new Workbook("report.xlsx");
Worksheet sheet = workbook.Worksheets[0];
Slicer slicer = sheet.Slicers[0];

SlicerCacheItemCollection items = slicer.SlicerCache.SlicerCacheItems;
foreach (SlicerCacheItem item in items)
    Console.WriteLine(item.Value + " selected=" + item.Selected);
```

### Pitfall - slicer does not filter the sheet until data is refreshed

Symptom: the slicer exists but filtering does nothing to the rows/totals you read.
Cause: a slicer over a pivot only filters the pivot cache; the pivot body (and any formulas over it) must be recalculated after selection changes. Aspose.Cells will not do that for you.
Fix: after changing `Selected`, refresh the pivot data and recalculate (see pivot-tables.md), then re-read the affected cells.

```csharp
var workbook = new Workbook("report.xlsx");
Worksheet sheet = workbook.Worksheets[0];
PivotTable pt = sheet.PivotTables[0];
Slicer slicer = sheet.Slicers[0];

slicer.SlicerCache.SlicerCacheItems[0].Selected = false;

pt.CalculateData(new PivotTableCalculateOption { RefreshData = true });
workbook.CalculateFormula();
// ... now read the filtered pivot cells
```

### Pitfall - slicer over a table needs the table, not raw cells

Symptom: `sheet.Slicers.Add(...)` throws or the slicer shows nothing.
Cause: slicers attach to a `ListObject` or a `PivotTable`, not to an arbitrary cell range. A range that is not a real Excel table has no slicer model.
Fix: create a `ListObject` first (see worksheets-rows-columns.md), or bind to a pivot instead.

## Sparklines

### Add a sparkline group

`SparklineGroups.Add(SparklineType, string dataRange, bool isVertical, CellArea locationRange)` returns an int index. The data range supplies one row/column of source values per sparkline; the location range is the block of cells that each hold one mini-chart (one cell per series).

```csharp
var workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];

for (int r = 0; r < 5; r++)
{
    sheet.Cells[r, 0].PutValue(r * 2);   // series 1 data, A1:A5
    sheet.Cells[r, 1].PutValue(r * 3);   // series 2 data, B1:B5
}

// location: D1 holds sparkline for A1:A5, E1 holds sparkline for B1:B5
var location = new CellArea { StartRow = 0, StartColumn = 3, EndRow = 0, EndColumn = 4 };
int idx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:B5", true, location);
SparklineGroup group = sheet.SparklineGroups[idx];
workbook.Save("out.xlsx", SaveFormat.Xlsx);
```

### Style and per-sparkline access

`SparklineGroup` owns the look for all sparklines in the group; `group.Sparklines` indexes the individual sparklines. Preset styles are `SparklinePresetStyleType.Style1`..`Style36` plus `Custom` - there are no named "Colorful/Dark/Light" members. Colors are set through each point-type's `CellsColor` (`HighPointColor`, `LowPointColor`, `NegativePointsColor`, `FirstPointColor`, `LastPointColor`, `SeriesColor`); assign `Color` on the `CellsColor`.

```csharp
var workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];
for (int r = 0; r < 5; r++) { sheet.Cells[r, 0].PutValue(r * 2); sheet.Cells[r, 1].PutValue(r * 3); }

var location = new CellArea { StartRow = 0, StartColumn = 3, EndRow = 0, EndColumn = 4 };
int idx = sheet.SparklineGroups.Add(SparklineType.Column, "A1:B5", true, location);
SparklineGroup group = sheet.SparklineGroups[idx];

group.PresetStyle = SparklinePresetStyleType.Style6;
group.ShowHighPoint = true;
group.HighPointColor.Color = Color.Red;
group.ShowLowPoint = true;
group.LowPointColor.Color = Color.Green;
group.ShowNegativePoints = true;
group.NegativePointsColor.Color = Color.Orange;
group.DisplayHidden = true;

foreach (Sparkline spark in group.Sparklines)
    Console.WriteLine(spark.DataRange);   // e.g. "Sheet1!$A$1:$A$5"
workbook.Save("out.xlsx", SaveFormat.Xlsx);
```

Note: `group.SparklineCollection` is obsolete in 26.x - use `group.Sparklines`.

### Pitfall - sparkline "shows nothing" in the saved file

Symptom: the sparkline renders as an empty cell, or the mini-chart has no line.
Cause: either the source range is empty, or the location range does not have one cell per data series (Excel requires a 1:1 mapping). A column sparkline with a single value also collapses to a thin line that is easy to miss.
Fix: double-check `CellArea` coordinates (they are 0-based), confirm the data cells actually contain numbers, and make `EndColumn - StartColumn + 1` (or rows) equal the number of series.

### Pitfall - a sparkline group for a hidden/blank row

Symptom: sparklines are missing for rows that have no data yet.
Cause: a sparkline is tied to exactly the source cells in its range; blank source cells are simply skipped, they are not "filled down".
Fix: add the sparkline group only after the data is present, or extend the data range to cover the full (possibly empty) range and let the blanks show as gaps.

## Related
- pivot-tables.md - the pivot a slicer filters; refresh before reading filtered output.
- worksheets-rows-columns.md - ListObject tables that slicers attach to.
- charts.md - Calculate() before rendering anything.
- styles-and-formatting.md - CellsColor-based colors used by sparkline points.
