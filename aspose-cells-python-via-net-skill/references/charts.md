# Charts

Read this when creating or customizing a chart, or when a rendered chart image or PDF is blank, stale, mislabeled, or wrongly sized.

Charts hang off `Worksheet.charts` (a `ChartCollection`). `Charts.add(...)` returns an int index; index the collection to get the `Chart`. Series live on `Chart.n_series` (a `SeriesCollection`); category (X-axis) labels are set once on `n_series.category_data`.

## Pitfalls

### Rendered chart image or PDF is blank, stale, or mislabeled

Symptom: `Chart.to_image` or `Chart.to_pdf` yields a blank or zero-sized image, shows pre-edit values, or labels land in the wrong place.

Cause: a chart's layout and cached series values are computed lazily, so rendering before that gives empty or stale output. For a chart bound to a pivot table the pivot cache may also be stale. Separately, fonts missing on the machine get substituted with different metrics, shifting label positions.

Fix: call `chart.calculate()` immediately before `to_image`/`to_pdf`. For a pivot chart call `chart.refresh_pivot_data()` first, then `calculate()`. If labels shift or change typeface, install the required fonts (see cross-platform-deployment.md for headless and container setups) or set a rendering fallback font (see conversion-and-rendering.md).

### Chart shows old numbers after you edit source cells

Symptom: after `put_value` on source cells, chart properties or the rendered picture still show the pre-edit values.

Cause: series values are cached from the last calculate and are not refreshed automatically.

Fix: call `chart.calculate()` before reading chart properties or rendering. `chart.is_chart_data_changed()` returns true when cells changed since the last calculate - use it to decide whether a re-render is needed.

### ToImage produces a zero-sized image for some chart types

Symptom: one specific chart renders to an empty or zero-sized image, or a blank PDF.

Cause: a few types are not supported for rendering: Surface3D, SurfaceWireframe3D, SurfaceContour, SurfaceContourWireframe, Bubble3D, and Map.

Fix: use a supported type. Column, Bar, Line, Pie, Scatter, Area, Doughnut, Radar, Bubble, stock, cylinder, cone, pyramid, and the modern Sunburst, Treemap, Waterfall, Funnel, Histogram, BoxWhisker, and ParetoLine all render.

### Series or category range is empty or rejected

Symptom: a series comes out empty when you pass a range such as "C3:A1".

Cause: cell ranges must run top-left to bottom-right. "A1:C3" is valid; "C3:A1" is not.

Fix: always order ranges top-left to bottom-right.

## Create a chart

`Charts.add(type, upper_left_row, upper_left_column, lower_right_row, lower_right_column)` - the last four arguments are 0-based cell anchors in row, column, row, column order (not pixels, not 1-based).

```python
import aspose.cells as gc
import aspose.cells.charts as ch

workbook = gc.Workbook()
sheet = workbook.worksheets[0]
sheet.cells.get("A1").put_value("Jan"); sheet.cells.get("A2").put_value("Feb")
sheet.cells.get("B1").put_value(120); sheet.cells.get("B2").put_value(150)
# anchors: rows 5..20, cols 0..8 (0-based)
idx = sheet.charts.add(ch.ChartType.COLUMN, 5, 0, 20, 8)
chart = sheet.charts[idx]
chart.n_series.add("B1:B2", True)     # True = data is column-oriented (vertical)
chart.n_series.category_data = "A1:A2"  # X-axis category labels
workbook.save("chart.xlsx", gc.SaveFormat.XLSX)
```

Or wire series and categories in one call with `set_chart_data_range` (pass the whole block; the second arg says the block is vertical, i.e. series in columns):

```python
import aspose.cells as gc
import aspose.cells.charts as ch

workbook = gc.Workbook()
sheet = workbook.worksheets[0]
for r in range(5):
    sheet.cells.get(r, 0).put_value("C" + str(r))
    sheet.cells.get(r, 1).put_value(r * 10)
chart = sheet.charts[sheet.charts.add(ch.ChartType.COLUMN, 6, 0, 20, 8)]
chart.set_chart_data_range("A1:B5", True)
workbook.save("chart2.xlsx", gc.SaveFormat.XLSX)
```

The `ChartType` enum covers every standard Excel type (column, bar, line, pie, scatter, area, doughnut, radar, bubble, stock, cylinder, cone, pyramid) plus modern types (sunburst, treemap, waterfall, funnel, histogram, box-whisker). Change one enum value to switch chart style.

## Position and size

The Add anchors set the initial box. Override exact pixels through the chart's shape (`Chart.chart_object`):

```python
import aspose.cells as gc
import aspose.cells.charts as ch

workbook = gc.Workbook()
sheet = workbook.worksheets[0]
chart = sheet.charts[sheet.charts.add(ch.ChartType.COLUMN, 5, 0, 20, 8)]
chart.chart_object.width = 640   # pixels
chart.chart_object.height = 400
chart.chart_object.x = 120       # pixel offset within the sheet
chart.chart_object.y = 80
```

## Customize title, axes, legend, series, labels, and color

Verify each member you set; all of the following are real. Title objects hang off the chart and each axis; set their `text`.

```python
import aspose.cells as gc
import aspose.cells.charts as ch
import aspose.pydrawing as drawing

workbook = gc.Workbook()
sheet = workbook.worksheets[0]
sheet.cells.get("A1").put_value("Q1"); sheet.cells.get("B1").put_value(120)
chart = sheet.charts[sheet.charts.add(ch.ChartType.COLUMN, 4, 0, 20, 8)]
s = chart.n_series.add("B1", True)
chart.n_series.category_data = "A1"
chart.title.text = "Revenue"
chart.category_axis.title.text = "Quarter"
chart.value_axis.title.text = "USD"
chart.legend.position = ch.LegendPositionType.BOTTOM  # TOP/BOTTOM/LEFT/RIGHT/CORNER
chart.n_series[s].name = "2026"
chart.n_series[s].data_labels.show_value = True      # also show_category_name
chart.n_series[s].area.foreground_color = drawing.Color.steel_blue
workbook.save("custom.xlsx", gc.SaveFormat.XLSX)
```

Hide the legend with `chart.show_legend = False`. Fix an axis scale with `chart.value_axis.is_automatic_max_value = False; chart.value_axis.max_value = 200;` (also `min_value`, `is_automatic_min_value`). Data-label number format follows Excel custom patterns via `data_labels.number_format` (for example `'"$"#,##0'`).

## Render to image

Always `calculate()` first. `ImageOrPrintOptions` controls resolution and format.

```python
import aspose.cells as gc
import aspose.cells.charts as ch
import aspose.cells.rendering as rd

workbook = gc.Workbook()
sheet = workbook.worksheets[0]
sheet.cells.get("A1").put_value("A"); sheet.cells.get("B1").put_value(10)
chart = sheet.charts[sheet.charts.add(ch.ChartType.COLUMN, 4, 0, 20, 8)]
chart.n_series.add("B1", True)
chart.calculate()  # compute layout + values before rendering
opts = rd.ImageOrPrintOptions()
opts.horizontal_resolution = 300
opts.vertical_resolution = 300
    # PNG is the default; for JPEG etc set opts.image_type = gc.drawing.ImageType.JPEG
chart.to_image("chart.png", opts)
```

`to_image` also has a `Stream` overload and a `to_image(string, ImageType)` overload for output without an options object. To test whether a re-render is needed after edits:

```python
import aspose.cells as gc
import aspose.cells.charts as ch

workbook = gc.Workbook()
sheet = workbook.worksheets[0]
sheet.cells.get("B1").put_value(10)
chart = sheet.charts[sheet.charts.add(ch.ChartType.COLUMN, 4, 0, 20, 8)]
chart.n_series.add("B1", True)
chart.calculate()
sheet.cells.get("B1").put_value(999)                 # edit source after calculate
if chart.is_chart_data_changed():
    chart.calculate()                                # -> True, so recompute
```

## Render to PDF

```python
import aspose.cells as gc
import aspose.cells.charts as ch

workbook = gc.Workbook()
sheet = workbook.worksheets[0]
sheet.cells.get("A1").put_value("A"); sheet.cells.get("B1").put_value(10)
chart = sheet.charts[sheet.charts.add(ch.ChartType.COLUMN, 4, 0, 20, 8)]
chart.n_series.add("B1", True)
chart.calculate()
chart.to_pdf("chart.pdf")
# fixed page (width, height in inches) + centered on the page:
chart.to_pdf("chart-7x7.pdf", 7.0, 7.0, gc.PageLayoutAlignmentType.CENTER, gc.PageLayoutAlignmentType.CENTER)
```

Both `to_pdf` signatures have `Stream` overloads for in-memory output.

## Pivot charts

Bind a chart to an existing pivot table through `Chart.pivot_source`, then refresh the pivot data before rendering. Build the pivot table first (see pivot-tables.md).

```python
import aspose.cells as gc
import aspose.cells.charts as ch

# Requires an existing pivot table on the sheet (see pivot-tables.md)
sheet = gc.Workbook("book-with-pivot.xlsx").worksheets[0]
chart = sheet.charts[sheet.charts.add(ch.ChartType.COLUMN, 20, 0, 40, 8)]
chart.pivot_source = sheet.name + "!PivotTable1"  # e.g. "Sheet1!PivotTable1"
chart.refresh_pivot_data()  # sync series from the pivot cache
chart.calculate()
chart.to_image("pivot-chart.png")
```

## Combo chart and secondary axis

Create one chart, then change a series' `type` and push it onto the secondary value axis:

```python
import aspose.cells as gc
import aspose.cells.charts as ch

workbook = gc.Workbook()
sheet = workbook.worksheets[0]
for r in range(4):
    sheet.cells.get(r, 0).put_value(1000 * (r + 1))
    sheet.cells.get(r, 1).put_value(r + 1)
chart = sheet.charts[sheet.charts.add(ch.ChartType.COLUMN, 5, 0, 20, 8)]
chart.n_series.add("A1:B4", True)
chart.n_series[1].type = ch.ChartType.LINE    # second series becomes a line
chart.n_series[1].plot_on_second_axis = True  # onto the secondary value axis
workbook.save("combo.xlsx", gc.SaveFormat.XLSX)
```

## Related
- pivot-tables.md - build and refresh the pivot behind a pivot chart.
- conversion-and-rendering.md - default_font and save/render options for output fidelity.
- cross-platform-deployment.md - fonts on Linux/Docker so chart text renders correctly.
- cell-values-and-data.md - writing the series source data.
- images-and-shapes.md - positioning pictures and other drawing objects.
