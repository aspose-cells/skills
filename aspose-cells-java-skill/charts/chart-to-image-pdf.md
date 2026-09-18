---
name: chart-to-image-pdf
description: Render an Aspose.Cells for Java chart to PNG / JPEG / PDF, including options for resolution, smoothing, transparency, and how to convert an entire sheet to image. Use when the user asks to export a chart as an image or PDF.
applies_to: Java
keywords: [chart to image, chart to PDF, render, ImageOrPrintOptions, SheetRender]
---

# Render chart to image or PDF

## Imports

```java
import com.aspose.cells.*;
```

## Chart → image

```java
Chart chart = ws.getCharts().get(0);
ImageOrPrintOptions options = new ImageOrPrintOptions();
options.setImageType(ImageType.PNG);
options.setHorizontalResolution(300);
options.setVerticalResolution(300);

chart.toImage("chart.png", options);
```

## Chart → multiple images (frames for animated charts)

```java
Chart chart = ws.getCharts().get(0);
ImageOrPrintOptions options = new ImageOrPrintOptions();
options.setImageType(ImageType.PNG);
chart.toImage("chart-%d.png", options);
```

Animated-bar / GIF charts produce one frame per index.

## Worksheet → image

```java
Chart chart = ws.getCharts().get(0);
ImageOrPrintOptions img = new ImageOrPrintOptions();
img.setImageType(ImageType.PNG);
img.setOnePagePerSheet(true);

SheetRender renderer = new SheetRender(ws, img);
for (int p = 0; p < renderer.getPageCount(); p++) {
    renderer.toImage(p, "page-" + p + ".png");
}
```

`SheetRender.getPageCount()` reports how many pages the print view
produces; iterate to write each.

## Worksheet → PDF (one PDF per sheet)

```java
PdfSaveOptions pdf = new PdfSaveOptions();
pdf.setOnePagePerSheet(true);
wb.save("report.pdf", pdf);
```

## Pitfalls

- **Resolution default is 96 DPI** — for print quality, set 300+ DPI.
- **JPEG chart images** lose crispness; use PNG for diagrams.
- **Translucent fills** may render solid black on white PDFs unless
  `setTransparent(true)` is set on `ImageOrPrintOptions`.
- **Multi-page worksheets** — `SheetRender` writes N pages; iterate
  with `toImage(page, stream)` rather than reading the
  first page only.
- **Charts embedded inside shapes** — `Chart.toImage(...)` only
  renders the chart object, not the shape decorations around it.

### Rendered chart image or PDF is blank, stale, or mislabeled

Symptom: `chart.toImage(...)` or `chart.toPdf(...)` yields a blank
or zero-sized image, shows pre-edit values, or labels land in the
wrong place.

Cause: a chart's layout and cached series values are computed
lazily, so rendering before that gives empty or stale output. For
a chart bound to a pivot table the pivot cache may also be stale.

Fix: call `chart.calculate()` immediately before `toImage` /
`toPdf`. For a pivot chart call `chart.refreshPivotData()` first,
then `calculate()`. If you might have edited the source cells
between renders, sanity-check with `chart.isChartDataChanged()` —
it returns `true` when cells changed since the last `calculate()`.

```java
Chart chart = ws.getCharts().get(0);
if (chart.isChartDataChanged()) {       // optional: only re-render when needed
    chart.calculate();                  // compute layout + values
}
ImageOrPrintOptions opts = new ImageOrPrintOptions();
opts.setImageType(ImageType.PNG);
opts.setHorizontalResolution(300);
opts.setVerticalResolution(300);
chart.toImage("chart.png", opts);
```

For a pivot-bound chart:

```java
Chart pivotChart = ws.getCharts().get(0);
pivotChart.refreshPivotData();           // sync series from the pivot cache
pivotChart.calculate();
pivotChart.toImage("pivot-chart.png");
```

### `toImage` produces a zero-sized image for some chart types

Symptom: one specific chart renders to an empty or zero-sized image,
or a blank PDF.

Cause: a few types are not supported for rendering: `SURFACE_3_D`,
`SURFACE_WIREFRAME_3_D`, `SURFACE_CONTOUR`,
`SURFACE_CONTOUR_WIREFRAME`, `BUBBLE_3_D`, and `MAP`.

Fix: use a supported type. Column, Bar, Line, Pie, Scatter, Area,
Doughnut, Radar, Bubble, stock, cylinder, cone, pyramid, and the
modern `SUNBURST`, `TREEMAP`, `WATERFALL`, `FUNNEL`, `HISTOGRAM`,
`BOX_WHISKER`, `PARETO_LINE` all render.

## Related

- [create-chart.md](create-chart.md)
- [../rendering/convert-to-pdf.md](../rendering/convert-to-pdf.md)
- [../rendering/worksheet-image.md](../rendering/worksheet-image.md)
- [../_shared/fonts-linux-docker.md](../_shared/fonts-linux-docker.md)
