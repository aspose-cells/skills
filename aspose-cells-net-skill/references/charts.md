# Charts

Read this when creating or customizing a chart, or when a rendered chart image or PDF is blank, stale, mislabeled, or wrongly sized.

Charts hang off `Worksheet.Charts` (a `ChartCollection`). `Charts.Add(...)` returns an int index; index the collection to get the `Chart`. Series live on `Chart.NSeries` (a `SeriesCollection`); category (X-axis) labels are set once on `NSeries.CategoryData`.

## Pitfalls

### Rendered chart image or PDF is blank, stale, or mislabeled

Symptom: `Chart.ToImage` or `Chart.ToPdf` yields a blank or zero-sized image, shows pre-edit values, or labels land in the wrong place.

Cause: a chart's layout and cached series values are computed lazily, so rendering before that gives empty or stale output. For a chart bound to a pivot table the pivot cache may also be stale. Separately, fonts missing on the machine get substituted with different metrics, shifting label positions.

Fix: call `chart.Calculate()` immediately before `ToImage`/`ToPdf`. For a pivot chart call `chart.RefreshPivotData()` first, then `Calculate()`. If labels shift or change typeface, install the required fonts (see cross-platform-deployment.md for headless and container setups) or set a rendering fallback font (see conversion-and-rendering.md).

### Chart shows old numbers after you edit source cells

Symptom: after `PutValue` on source cells, chart properties or the rendered picture still show the pre-edit values.

Cause: series values are cached from the last calculate and are not refreshed automatically.

Fix: call `chart.Calculate()` before reading chart properties or rendering. `chart.IsChartDataChanged()` returns true when cells changed since the last calculate - use it to decide whether a re-render is needed.

### ToImage produces a zero-sized image for some chart types

Symptom: one specific chart renders to an empty or zero-sized image, or a blank PDF.

Cause: a few types are not supported for rendering: Surface3D, SurfaceWireframe3D, SurfaceContour, SurfaceContourWireframe, Bubble3D, and Map.

Fix: use a supported type. Column, Bar, Line, Pie, Scatter, Area, Doughnut, Radar, Bubble, stock, cylinder, cone, pyramid, and the modern Sunburst, Treemap, Waterfall, Funnel, Histogram, BoxWhisker, and ParetoLine all render.

### Series or category range is empty or rejected

Symptom: a series comes out empty when you pass a range such as "C3:A1".

Cause: cell ranges must run top-left to bottom-right. "A1:C3" is valid; "C3:A1" is not.

Fix: always order ranges top-left to bottom-right.

## Create a chart

`Charts.Add(type, upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn)` - the last four arguments are 0-based cell anchors in row, column, row, column order (not pixels, not 1-based).

```csharp
var workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];
sheet.Cells["A1"].PutValue("Jan"); sheet.Cells["A2"].PutValue("Feb");
sheet.Cells["B1"].PutValue(120);   sheet.Cells["B2"].PutValue(150);
// anchors: rows 5..20, cols 0..8 (0-based)
int idx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
Chart chart = sheet.Charts[idx];
chart.NSeries.Add("B1:B2", true);     // true = data is column-oriented (vertical)
chart.NSeries.CategoryData = "A1:A2"; // X-axis category labels
workbook.Save("chart.xlsx", SaveFormat.Xlsx);
```

Or wire series and categories in one call with `SetChartDataRange` (pass the whole block; the second arg says the block is vertical, i.e. series in columns):

```csharp
var workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];
for (int r = 0; r < 5; r++) { sheet.Cells[r, 0].PutValue("C" + r); sheet.Cells[r, 1].PutValue(r * 10); }
Chart chart = sheet.Charts[sheet.Charts.Add(ChartType.Column, 6, 0, 20, 8)];
chart.SetChartDataRange("A1:B5", true);
workbook.Save("chart2.xlsx", SaveFormat.Xlsx);
```

The `ChartType` enum covers every standard Excel type (column, bar, line, pie, scatter, area, doughnut, radar, bubble, stock, cylinder, cone, pyramid) plus modern types (sunburst, treemap, waterfall, funnel, histogram, box-whisker). Change one enum value to switch chart style.

## Position and size

The Add anchors set the initial box. Override exact pixels through the chart's shape (`Chart.ChartObject`):

```csharp
var workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];
Chart chart = sheet.Charts[sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8)];
chart.ChartObject.Width = 640;   // pixels
chart.ChartObject.Height = 400;
chart.ChartObject.X = 120;       // pixel offset within the sheet
chart.ChartObject.Y = 80;
```

## Customize title, axes, legend, series, labels, and color

Verify each member you set; all of the following are real. Title objects hang off the chart and each axis; set their `Text`.

```csharp
var workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];
sheet.Cells["A1"].PutValue("Q1"); sheet.Cells["B1"].PutValue(120);
Chart chart = sheet.Charts[sheet.Charts.Add(ChartType.Column, 4, 0, 20, 8)];
int s = chart.NSeries.Add("B1", true);
chart.NSeries.CategoryData = "A1";
chart.Title.Text = "Revenue";
chart.CategoryAxis.Title.Text = "Quarter";
chart.ValueAxis.Title.Text = "USD";
chart.Legend.Position = LegendPositionType.Bottom; // Top/Bottom/Left/Right/Corner
chart.NSeries[s].Name = "2026";
chart.NSeries[s].DataLabels.ShowValue = true;      // also ShowCategoryName
chart.NSeries[s].Area.ForegroundColor = Color.SteelBlue;
workbook.Save("custom.xlsx", SaveFormat.Xlsx);
```

Hide the legend with `chart.ShowLegend = false`. Fix an axis scale with `chart.ValueAxis.IsAutomaticMaxValue = false; chart.ValueAxis.MaxValue = 200;` (also `MinValue`, `IsAutomaticMinValue`). Data-label number format follows Excel custom patterns via `DataLabels.NumberFormat` (for example `"\"$\"#,##0"`).

## Render to image

Always `Calculate()` first. `ImageOrPrintOptions` controls resolution and format.

```csharp
var workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];
sheet.Cells["A1"].PutValue("A"); sheet.Cells["B1"].PutValue(10);
Chart chart = sheet.Charts[sheet.Charts.Add(ChartType.Column, 4, 0, 20, 8)];
chart.NSeries.Add("B1", true);
chart.Calculate(); // compute layout + values before rendering
var opts = new ImageOrPrintOptions { HorizontalResolution = 300, VerticalResolution = 300 };
// PNG is the default; for JPEG etc set opts.ImageType = Aspose.Cells.Drawing.ImageType.Jpeg
chart.ToImage("chart.png", opts);
```

`ToImage` also has `Stream` overloads and a `ToImage(string, Aspose.Cells.Drawing.ImageType)` overload for output without an options object. To test whether a re-render is needed after edits:

```csharp
var workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];
sheet.Cells["B1"].PutValue(10);
Chart chart = sheet.Charts[sheet.Charts.Add(ChartType.Column, 4, 0, 20, 8)];
chart.NSeries.Add("B1", true);
chart.Calculate();
sheet.Cells["B1"].PutValue(999);                    // edit source after calculate
if (chart.IsChartDataChanged()) chart.Calculate();  // -> true, so recompute
```

## Render to PDF

```csharp
var workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];
sheet.Cells["A1"].PutValue("A"); sheet.Cells["B1"].PutValue(10);
Chart chart = sheet.Charts[sheet.Charts.Add(ChartType.Column, 4, 0, 20, 8)];
chart.NSeries.Add("B1", true);
chart.Calculate();
chart.ToPdf("chart.pdf");
// fixed page (width, height in inches) + centered on the page:
chart.ToPdf("chart-7x7.pdf", 7f, 7f, PageLayoutAlignmentType.Center, PageLayoutAlignmentType.Center);
```

Both `ToPdf` signatures have `Stream` overloads for in-memory output.

## Pivot charts

Bind a chart to an existing pivot table through `Chart.PivotSource`, then refresh the pivot data before rendering. Build the pivot table first (see pivot-tables.md).

```csharp
// Requires an existing pivot table on the sheet (see pivot-tables.md)
Worksheet sheet = new Workbook("book-with-pivot.xlsx").Worksheets[0];
Chart chart = sheet.Charts[sheet.Charts.Add(ChartType.Column, 20, 0, 40, 8)];
chart.PivotSource = sheet.Name + "!PivotTable1"; // e.g. "Sheet1!PivotTable1"
chart.RefreshPivotData();  // sync series from the pivot cache
chart.Calculate();
chart.ToImage("pivot-chart.png");
```

## Combo chart and secondary axis

Create one chart, then change a series' `Type` and push it onto the secondary value axis:

```csharp
var workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];
for (int r = 0; r < 4; r++) { sheet.Cells[r, 0].PutValue(1000 * (r + 1)); sheet.Cells[r, 1].PutValue(r + 1); }
Chart chart = sheet.Charts[sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8)];
chart.NSeries.Add("A1:B4", true);
chart.NSeries[1].Type = ChartType.Line;   // second series becomes a line
chart.NSeries[1].PlotOnSecondAxis = true; // onto the secondary value axis
workbook.Save("combo.xlsx", SaveFormat.Xlsx);
```

## Related
- pivot-tables.md - build and refresh the pivot behind a pivot chart.
- conversion-and-rendering.md - DefaultFont and save/render options for output fidelity.
- cross-platform-deployment.md - fonts on Linux/Docker so chart text renders correctly.
- cell-values-and-data.md - writing the series source data.
- images-and-shapes.md - positioning pictures and other drawing objects.
