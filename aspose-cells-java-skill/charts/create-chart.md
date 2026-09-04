---
name: create-chart
description: Create a chart on a worksheet in Aspose.Cells for Java — pick a ChartType, anchor a chart to a range, set the data source and titles. Use when the user wants any kind of chart on a sheet.
applies_to: Java
keywords: [chart, addChart, COLUMNN, PIE, LINE, BAR, area, scatter, series, axis]
---

# Create a chart

## Imports

```java
import com.aspose.cells.*;
```

## Quick example — clustered column chart

```java
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().get(0);

// Sample data
ws.getCells().get("A1").setValue("Region");
ws.getCells().get("B1").setValue("Sales");
ws.getCells().get("A2").setValue("East");  ws.getCells().get("B2").setValue(1234);
ws.getCells().get("A3").setValue("West");  ws.getCells().get("B3").setValue(2500);
ws.getCells().get("A4").setValue("North"); ws.getCells().get("B4").setValue(1700);

// Anchor: row 5 col 0 to row 22 col 8
int idx = ws.getCharts().add(ChartType.COLUMN, 5, 0, 22, 8);
Chart chart = ws.getCharts().get(idx);

// Data range includes the header so column names become series names
chart.setChartDataRange("A1:B4", true);

// Surface to user
chart.getTitle().setText("Quarterly Sales");
chart.getTitle().setVisible(true);

wb.save("column-chart.xlsx");
```

The boolean `true` passed to `setChartDataRange` means "first row holds
the category labels". Pass `false` if the first row is data.

## Other common chart types

```java
ws.getCharts().add(ChartType.PIE,     0, 0, 15, 8);   // pie
ws.getCharts().add(ChartType.LINE,    0, 0, 15, 8);   // line
ws.getCharts().add(ChartType.BAR,     0, 0, 15, 8);   // horizontal bar
ws.getCharts().add(ChartType.SCATTER, 0, 0, 15, 8);   // X-Y
ws.getCharts().add(ChartType.AREA,    0, 0, 15, 8);   // area
ws.getCharts().add(ChartType.DOUGHNUT,0, 0, 15, 8);   // doughnut
ws.getCharts().add(ChartType.RADAR,   0, 0, 15, 8);   // radar
ws.getCharts().add(ChartType.BUBBLE,  0, 0, 15, 8);   // bubble
```

For 3D variants append `3D` (e.g. `COLUMN_3_D`, `LINE_3_D`, `PIE_3_D`); for
stacked / 100 % stacked use `*Stacked` / `*100PercentStacked`.

## Show / hide parts

```java
import static com.aspose.cells.Color.fromArgb;
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().get(0);
int idx = ws.getCharts().add(ChartType.COLUMN, 0, 0, 15, 8);
Chart chart = ws.getCharts().get(idx);
chart.getTitle().setText("Q1 Report");
chart.getLegend().setPosition(LegendPositionType.BOTTOM);
chart.getPlotArea().getArea().setForegroundColor(fromArgb(0xFFFFFF)); // white
```

`Chart.setPlotVisibleCells(false)` hides a chart; there is no
`setHidden(boolean)` — control visibility through that or by removing
the chart from the collection. The `import static` disambiguates `Color`
from the prelude's `java.awt.Color`.

## Pitfalls

- **Range order matters** — `"A1:B4"` and not `"B4:A1"`. The
  row/column ordering rule is the same as for pivot tables.
- **`setChartDataRange(range, true)`** expects headers in the first
  row. If your data has only numbers, pass `false`; otherwise you
  lose the first row of data.
- **Adding a chart twice** at the same anchor produces two chart
  frames overlapping.
- **Re-using `Chart` from previous tests** — chart instances belong
  to a worksheet; cloning requires recreating the chart.
- **Don't forget to save** — a chart added in memory does not persist
  until `wb.save(...)` is called.

## Related

- [chart-types-selection.md](chart-types-selection.md)
- [chart-data-formatting.md](chart-data-formatting.md)
- [chart-to-image-pdf.md](chart-to-image-pdf.md)
