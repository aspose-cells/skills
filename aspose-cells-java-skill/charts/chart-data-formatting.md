---
name: chart-data-formatting
description: Configure chart series, axes, data labels, legends, plot area, trendlines, and conditional series colours in Aspose.Cells for Java. Use when the user asks to colour bars, set axis units, or add data labels.
applies_to: Java
keywords: [series, axis, data labels, legend, trendline, colors, plot area]
---

# Chart data & formatting

## Imports

```java
import com.aspose.cells.*;
```

## Add / configure series

```java
Chart chart = ws.getCharts().get(0);

// Existing series
Series s = chart.getNSeries().get(0);
s.setName("Series A");
s.setValues("=Sheet1!$B$2:$B$10");
s.setXValues("=Sheet1!$A$2:$A$10");

// Add a new series; add(...) returns the index, not the Series
int idx = chart.getNSeries().add("=Sheet1!$C$2:$C$10", true);
Series s2 = chart.getNSeries().get(idx);
```

## Series colours

```java
Chart chart = ws.getCharts().get(0);
Series series = chart.getNSeries().get(0);
series.getArea().setForegroundColor(com.aspose.cells.Color.fromArgb(0xFFA500));   // orange
series.getBorder().setColor(com.aspose.cells.Color.fromArgb(0x404040));         // dark gray

ChartPointCollection points = series.getPoints();
for (int i = 0; i < points.getCount(); i++) {
    points.get(i).getArea().setForegroundColor(com.aspose.cells.Color.fromArgb(0xADD8E6)); // light blue
}
```

## Data labels

```java
Chart chart = ws.getCharts().get(0);
DataLabels labels = chart.getNSeries().get(0).getDataLabels();
labels.setShowValue(true);
labels.setPosition(LabelPositionType.OUTSIDE_END);
labels.setNumberFormat("0.0\"K\"");
labels.getFont().setSize(9);
```

## Axes

```java
Chart chart = ws.getCharts().get(0);
chart.getCategoryAxis().getTickLabels().getFont().setSize(8);
chart.getCategoryAxis().getTitle().setText("Month");
chart.getValueAxis().setMaxValue(50_000.0);
chart.getValueAxis().setMinValue(0.0);
chart.getValueAxis().setMajorUnit(10_000.0);
```

## Legend

```java
Chart chart = ws.getCharts().get(0);
chart.getLegend().setPosition(LegendPositionType.BOTTOM);
// LegendPositionType.BOTTOM already places the legend outside the plot area.
```

## Trendline

```java
Chart chart = ws.getCharts().get(0);
TrendlineCollection t = chart.getNSeries().get(0).getTrendLines();
int idx = t.add(TrendlineType.LINEAR);
Trendline line = t.get(idx);
line.setDisplayRSquared(true);
line.setDisplayEquation(true);
```

## Pitfalls

- **`Series.setName(...)`** — pass a string literal **or** a cell
  reference like `"=Sheet1!$B$1"`; passing both the literal and a
  separate `setValues` does not change the display name.
- **DataLabels conditional on a series** — use `getPoints().get(i)
  .getDataLabels()` to override per-point; the series-level settings
  apply only when no point override exists.
- **Axis min/max conflict with log scale** — when
  `setLogarithmic(true)`, min must be > 0.
- **Series-add boolean** — when adding a series via the
  `SeriesCollection.add(String, boolean)` overload, the trailing
  boolean is `isVertical`; `true` reads the range column-by-column,
  `false` row-by-row. There is no `setValues(..., boolean)` overload —
  `Series.setValues(String)` takes only the range string.
- **Re-using `series.getPoints()`** — the iterator may be invalidated
  if you mutate series data inside the loop.

## Related

- [create-chart.md](create-chart.md)
- [chart-to-image-pdf.md](chart-to-image-pdf.md)
