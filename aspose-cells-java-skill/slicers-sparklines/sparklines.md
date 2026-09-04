---
name: sparklines
description: Add and style sparklines (mini line/column/win-loss charts inside a cell) with Aspose.Cells for Java. Use when the user wants an in-cell trend or distribution visualization, or to set high/low/first/last/negative point colors and preset styles.
applies_to: Java
keywords: [sparkline, SparklineGroup, mini chart, in-cell, CellsColor, preset style, trendline]
---

# Sparklines

Sparklines are tiny in-cell charts (line, column, stacked, or
win/loss). They live on `Worksheet.getSparklineGroups()` (a
`SparklineGroupCollection`). One `SparklineGroup` covers one source
data range and a contiguous block of location cells; each location cell
holds one mini-chart.

## Imports

```java
import com.aspose.cells.*;
```

## Add a sparkline group

`SparklineGroupCollection.add(type, dataRange, isVertical, locationArea)`
returns an int index. The data range supplies one row/column of source
values per sparkline; the location `CellArea` is the block of cells
that each hold one mini-chart.

```java
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().get(0);

// Source data — A1:A5 (column 0) and B1:B5 (column 1).
for (int r = 0; r < 5; r++) {
    ws.getCells().get(r, 0).setValue(r * 2);
    ws.getCells().get(r, 1).setValue(r * 3);
}

// Location: D1 holds the sparkline for A1:A5, E1 for B1:B5.
CellArea location = CellArea.createCellArea(0, 3, 0, 4);
int idx = ws.getSparklineGroups().add(
    SparklineType.LINE, "A1:B5", true, location);
SparklineGroup group = ws.getSparklineGroups().get(idx);

wb.save("out.xlsx");
```

`SparklineType` is `LINE`, `COLUMN`, `STACKED`, or `WIN_LOSS`.

## Style and per-sparkline access

`SparklineGroup` owns the look for all sparklines in the group;
`group.getSparklines()` returns the `SparklineCollection`. Preset styles
are constants on `SparklinePresetStyleType` (`STYLE_1` … `STYLE_36`,
`CUSTOM`) — there are no named "Colorful/Dark/Light" members. Colors
are set through each point-type's `CellsColor` (`HighPointColor`,
`LowPointColor`, `NegativePointsColor`, `FirstPointColor`,
`LastPointColor`, `MarkersColor`, `SeriesColor`); assign `Color` on the
`CellsColor`.

```java
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().get(0);
for (int r = 0; r < 5; r++) {
    ws.getCells().get(r, 0).setValue(r * 2);
    ws.getCells().get(r, 1).setValue(r * 3);
}

CellArea location = CellArea.createCellArea(0, 3, 0, 4);
int idx = ws.getSparklineGroups().add(
    SparklineType.COLUMN, "A1:B5", true, location);
SparklineGroup group = ws.getSparklineGroups().get(idx);

group.setPresetStyle(SparklinePresetStyleType.STYLE_6);
group.setShowHighPoint(true);
group.getHighPointColor().setColor(com.aspose.cells.Color.getRed());
group.setShowLowPoint(true);
group.getLowPointColor().setColor(com.aspose.cells.Color.getGreen());
group.setShowNegativePoints(true);
group.getNegativePointsColor().setColor(com.aspose.cells.Color.getOrange());
group.setShowFirstPoint(true);
group.getFirstPointColor().setColor(com.aspose.cells.Color.getBlue());
group.setShowLastPoint(true);
group.getLastPointColor().setColor(com.aspose.cells.Color.getPurple());
group.setDisplayHidden(true);                      // plot hidden source cells too

for (int i = 0; i < group.getSparklines().getCount(); i++) {
    Sparkline spark = group.getSparklines().get(i);
    System.out.println(spark.getDataRange());       // e.g. "Sheet1!A1:A5"
}
wb.save("out.xlsx");
```

`SparklineGroup.getSparklines()` is the supported accessor — the older
`getSparklineCollection()` exists but is the legacy alias; prefer
`getSparklines()`.

## Read a sparkline's data range and position

Each `Sparkline` exposes `getDataRange()`, `getRow()`, and
`getColumn()`. `getRow()` and `getColumn()` are 0-based and identify
the location cell holding this mini-chart.

## Render a single sparkline to image

`Sparkline.toImage(OutputStream, ImageOrPrintOptions)` renders just
that one sparkline as a PNG/JPEG/etc. stream. Use this when you want a
separate image per cell rather than the surrounding workbook export.

```java
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().get(0);
for (int r = 0; r < 5; r++) {
    ws.getCells().get(r, 0).setValue(r * 2);
}
CellArea location = CellArea.createCellArea(0, 1, 0, 1);
int idx = ws.getSparklineGroups().add(
    SparklineType.LINE, "A1:A5", false, location);
SparklineGroup group = ws.getSparklineGroups().get(idx);

Sparkline first = group.getSparklines().get(0);
ImageOrPrintOptions opts = new ImageOrPrintOptions();
opts.setImageType(ImageType.PNG);
first.toImage("spark.png", opts);
```

## Pitfalls

### Sparkline "shows nothing" in the saved file

Symptom: the sparkline renders as an empty cell, or the mini-chart has
no line.
Cause: either the source range is empty, or the location range does
not have one cell per data series (Excel requires a 1:1 mapping). A
column sparkline with a single value also collapses to a thin line
that is easy to miss.
Fix: confirm `CellArea` coordinates are 0-based (`StartRow`, `StartColumn`,
`EndRow`, `EndColumn`), that the data cells actually contain numbers,
and that `EndColumn - StartColumn + 1` (or rows) equals the number of
series.

### Sparkline group for a hidden/blank row

Symptom: sparklines are missing for rows that have no data yet.
Cause: a sparkline is tied to exactly the source cells in its range;
blank source cells are simply skipped, they are not "filled down".
Fix: add the sparkline group only after the data is present, or extend
the data range to cover the full (possibly empty) range and let the
blanks show as gaps. Set `setDisplayHidden(true)` on the group if you
want to plot hidden-source-cell values.

### `CellsColor.setColor(...)` takes `com.aspose.cells.Color`, not `java.awt.Color`

Symptom: snippet compiles only after a static import, then `Color` is
ambiguous with `java.awt.Color`.
Cause: the wrapper already imports `java.awt.Color`. A bare `Color`
in the snippet is ambiguous.
Fix: fully qualify the Aspose `Color` type when assigning to
`CellsColor.setColor(...)` — for example, `Color.getRed()` from the
`com.aspose.cells` package.

## Resetting a group's ranges

If the data range or location range needs to change after creation,
call `SparklineGroup.resetRanges(dataRange, isVertical, locationArea)`
rather than adding a new group. `clearSparklines(CellArea)` removes
just the sparklines in a given area; `clearSparklineGroups(CellArea)`
removes entire groups that overlap the area.

## Notes

- `SparklineGroup` colors accept `CellsColor`. Use a solid RGB getter
  from the Aspose `Color` type (for example,
  `setColor(com.aspose.cells.Color.getRed())`) for solid RGB;
  `setArgb(int)` for an explicit ARGB int;
  `setThemeColor(ThemeColor)` / `setColorIndex(int)` for theme colors.
- For the full sparkline group API see the `SparklineGroup` class —
  there are also `setShowMarkers(boolean)`, `getMarkersColor()`, axis
  controls (`setVerticalAxisMaxValueType`,
  `setVerticalAxisMaxValue(int, double)`, etc.), and
  `setPlotRightToLeft(boolean)`.

## Related

- [slicers.md](slicers.md) — interactive filter buttons, often shown
  alongside sparklines in dashboards.
- [../charts/chart-data-formatting.md](../charts/chart-data-formatting.md) —
  regular charts and their color formatting.
- [../cells/cell-formatting-styles.md](../cells/cell-formatting-styles.md) —
  `CellsColor` and theme colors.
