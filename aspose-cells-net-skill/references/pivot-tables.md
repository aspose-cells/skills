# Pivot Tables

Read this when creating a pivot table, changing its source range or fields, or when a
pivot shows no data, only headers, or stale (previous-run) numbers.

Aspose.Cells writes the pivot *definition* plus a cached snapshot of values; it never
recomputes the body automatically. Every value in the pivot area is stale until you
calculate. `CalculateData()` computes the view from the inner *pivot cache*; it does not
re-read the source range. If the source data changed, refresh the cache first.

## Pitfalls

### Pivot body is empty, shows only headers, or shows stale numbers

Symptom: after `PivotTables.Add(...)`, or after editing the source cells, the pivot area is
blank, shows field headers with no values, or keeps the numbers from a previous run.
Cause: the body is computed on demand from the pivot cache. `Add` builds the cache once
from the source at that moment; later edits to the source cells do not flow into the cache,
and nothing recomputes on save.
Fix, in order:
1. If the source cells are formula-driven, call `workbook.CalculateFormula()` first so the
   source values are current (Aspose does not auto-evaluate formulas either).
2. Rebuild the cache from the source and recompute the view in one non-obsolete call:
   `pt.CalculateData(new PivotTableCalculateOption { RefreshData = true });`.
3. Do this before `Save()` and before reading any pivot output cell.

```csharp
var workbook = new Workbook("report.xlsx");
PivotTable pt = workbook.Worksheets[0].PivotTables[0];
workbook.CalculateFormula();          // only needed if the source cells are formulas
// rebuild cache from the current source, then recompute the pivot view:
pt.CalculateData(new PivotTableCalculateOption { RefreshData = true });
workbook.Save("report.xlsx");
```

`pt.CalculateData()` with no option recomputes only from the *existing* cache. That is fine
immediately after `Add` (the cache was just built), but it will NOT pick up later edits to
the source cells.

To refresh the cache from the current source, then recompute the view, use the one-call
form `pt.CalculateData(new PivotTableCalculateOption { RefreshData = true })`. The
two-step equivalent `pt.PivotCache.Refresh(); pt.CalculateData();` also works -
`PivotCache` is exposed as `PivotTable.PivotCache` in 26.7+. `CalculateData()` itself is never obsolete - only the *refresh* step changed: the old
`pt.RefreshData()` is obsolete in 26.7 (marked obsolete - use `pt.PivotCache.Refresh()`
instead); it still runs but emits a warning. So `pt.RefreshData(); pt.CalculateData();`
works, but prefer the non-obsolete `pt.PivotCache.Refresh(); pt.CalculateData();` or the
one-call `pt.CalculateData(new PivotTableCalculateOption { RefreshData = true })`.

### Cannot compute correctly in code - defer refresh to Excel

Symptom: the pivot depends on a source you cannot fully reproduce headless, or you want the
numbers to be live when the user opens the file.
Cause: you do not want to bake a computed snapshot at save time.
Fix: set `pt.RefreshDataOnOpeningFile = true;`. Excel then refreshes the pivot when the file
is opened. Note this does not populate the body inside the saved file, so a headless reader
(including Aspose re-opening it) still sees the old cache until it calculates.

### Reading pivot output cells returns null before calculation

Symptom: reading cells in the pivot area returns null/empty even though the pivot looks set
up.
Cause: pivot output cells are ordinary cells only *after* the body is computed.
Fix: call `CalculateData` (see above), then read via the pivot's ranges. `DataBodyRange` is
the values region; `TableRange1` is the whole report excluding page fields; `TableRange2`
includes page fields. All are `CellArea` with `StartRow`/`StartColumn`/`EndRow`/`EndColumn`.

```csharp
var workbook = new Workbook("report.xlsx");
Worksheet sheet = workbook.Worksheets[0];
PivotTable pt = sheet.PivotTables[0];
pt.CalculateData();                                  // body cells are null before this
CellArea body = pt.DataBodyRange;                    // values region
object v = sheet.Cells[body.StartRow, body.StartColumn].Value;
```

### Recomputing stale output after source edits

Symptom: the pivot still shows old numbers after you change the source cells.
Cause: the pivot body is a snapshot computed from the cache; editing the source does not
recompute it.
Fix: call `CalculateData(new PivotTableCalculateOption { RefreshData = true })` (or the
two-step `pt.PivotCache.Refresh(); CalculateData();`) after source edits and before
reading/saving. `PivotCache` is public in 26.7+ (`pt.PivotCache.Refresh()`), and
`GetDependentPivotTables()` exists for child pivots - use those; the members above are the
recommended path.

## Create a pivot table

`PivotTables.Add(source, destCell, name)` returns the new pivot's index. `source` is a
range string like `"=Sheet1!A1:C100"` (top-left to bottom-right; reversed ranges are
invalid) or a named range. Add fields by their source column index into an area with
`AddFieldToArea(PivotFieldType, columnIndex)`.

```csharp
Worksheet sheet = new Workbook().Worksheets[0];
// source data assumed to live in Sheet1!A1:C100 with headers in row 1
int idx = sheet.PivotTables.Add("=Sheet1!A1:C100", "F3", "SalesPivot");
PivotTable pt = sheet.PivotTables[idx];
pt.AddFieldToArea(PivotFieldType.Row, 0);      // 1st source column -> Row area
pt.AddFieldToArea(PivotFieldType.Column, 1);   // 2nd column -> Column area
pt.AddFieldToArea(PivotFieldType.Data, 2);     // 3rd column -> Data (values) area
pt.CalculateData();                            // compute the body before save/read
```

`PivotFieldType` values: `Row`, `Column`, `Data`, `Page`, `Undefined`. `AddFieldToArea`
also has overloads taking a field name (`string`) or a `PivotField`. There are `Add`
overloads with an `int row, int column` destination instead of a cell string.

## Aggregation function and number format on a data field

`DataFields` is a collection; each entry is a `PivotField`. Set its `Function`
(`ConsolidationFunction`) and `NumberFormat` (an Excel format string).

```csharp
PivotTable pt = new Workbook("data.xlsx").Worksheets[0].PivotTables[0];
PivotField data0 = pt.DataFields[0];
data0.Function = ConsolidationFunction.Sum;    // Average, Count, Max, Min, Product, ...
data0.NumberFormat = "#,##0.00";
data0.DisplayName = "Total Sales";
```

`ConsolidationFunction` includes `Sum`, `Average`, `Count`, `CountNums`, `DistinctCount`,
`Max`, `Min`, `Product`, `StdDev`, `StdDevp`, `Var`, `Varp` (`DistinctCount` needs Excel
2013+).

## Page filter, grand totals, style, empty-cell text

```csharp
PivotTable pt = new Workbook("data.xlsx").Worksheets[0].PivotTables[0];
pt.AddFieldToArea(PivotFieldType.Page, 3);     // report (page) filter field
pt.ShowRowGrandTotals = true;                  // ShowRowGrandTotals/ColumnGrandTotals preferred
pt.ShowColumnGrandTotals = true;
pt.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
pt.DisplayNullString = true;
pt.NullString = "-";                           // "for empty cells show" text
```

Layout form: `pt.ShowInCompactForm()`, `pt.ShowInOutlineForm()`, `pt.ShowInTabularForm()`.
Style names run `PivotTableStyleLight1..28`, `...Medium1..28`, `...Dark1..28`, plus `None`.

## Manual cell formatting

`FormatAll(style)` styles the whole pivot; `Format(row, col, style)` styles one cell. Build
the `Style` from the workbook, never with a bare constructor.

```csharp
var workbook = new Workbook("data.xlsx");
PivotTable pt = workbook.Worksheets[0].PivotTables[0];
Style style = workbook.CreateStyle();
style.ForegroundColor = Color.LightYellow;
style.Pattern = BackgroundType.Solid;
pt.FormatAll(style);                           // Format(r, c, style) for a single cell
```

## Calculated field

`AddCalculatedField(name, formula, dragToDataArea)` adds a field computed from other fields;
pass `true` to drop it into the Data area immediately. Recompute afterward.

```csharp
PivotTable pt = new Workbook("data.xlsx").Worksheets[0].PivotTables[0];
pt.AddCalculatedField("Margin", "=Revenue-Cost", true);   // true -> into Data area
pt.CalculateData(new PivotTableCalculateOption { RefreshData = true });
```

## Change the data source

`ChangeDataSource(string[])` repoints the pivot; pass a one-element array for a single
range. Recompute after changing it.

```csharp
PivotTable pt = new Workbook("data.xlsx").Worksheets[0].PivotTables[0];
pt.ChangeDataSource(new string[] { "=Data!A1:D500" });    // top-left to bottom-right
pt.CalculateData(new PivotTableCalculateOption { RefreshData = true });
```

## Refresh nested / child pivots

A pivot whose source is another pivot is a child. Refresh the parent first, then each child
returned by `GetDependentPivotTables()` (the old `GetChildren()` is obsolete in 26.7).

```csharp
var workbook = new Workbook("nested.xlsx");
PivotTable parent = workbook.Worksheets[0].PivotTables[0];
var opt = new PivotTableCalculateOption { RefreshData = true };
parent.CalculateData(opt);                                // parent cache first
foreach (PivotTable child in parent.GetDependentPivotTables())
    child.CalculateData(opt);                             // then each child
```

## Sort, filter, delete

Auto-sort a field on its values: `PivotField f = pt.RowFields[0]; f.IsAutoSort = true;
f.IsAscendSort = true;`. Clear all field filters with `pt.ClearFilters()`. Delete a pivot
with `sheet.PivotTables.RemoveAt(index)` or `sheet.PivotTables.Remove(pt)`.

## Pivot charts

Bind a chart to a pivot via the chart's pivot source, then refresh the pivot and the chart
before rendering. See charts.md.

## Related
- formulas-and-calculation.md
- charts.md
- cell-values-and-data.md
- styles-and-formatting.md
- worksheets-rows-columns.md
