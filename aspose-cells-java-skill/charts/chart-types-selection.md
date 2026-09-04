---
name: chart-types-selection
description: Pick the right ChartType for the user's data — column, bar, line, pie, scatter, area, surface, radar, combo. Use when the user has data and wants to know what chart makes sense, or asks to compare types.
applies_to: Java
keywords: [chart type, ChartType, column, line, pie, scatter, area, combo, stock]
---

# Choose a chart type

The `ChartType` enum holds ~80 variants. Pick by **the question the
chart must answer**:

## Decision table

| User wants to show                  | ChartType                |
| ----------------------------------- | ------------------------ |
| Counts or values across categories  | `COLUMN`                 |
| Compare magnitudes of categories    | `BAR`                    |
| Trend over time                     | `LINE` (with markers)    |
| Share of total                      | `PIE` / `DOUGHNUT`       |
| Distribution or comparison of parts  | `COLUMN_STACKED`         |
| Correlation between two series       | `SCATTER`                |
| Three measures                      | `BUBBLE`                 |
| Distributions or fill              | `AREA`                   |
| Multivariate measurement            | `RADAR`                  |
| Surface / heatmap                   | `SURFACE_3_D`            |
| Stock OHLC                          | `STOCK_HIGH_LOW_CLOSE` etc. |

Combo charts (column + line on the same axes) and Excel-2016 types
(`BOX_WHISKER`, `FUNNEL`, `WATERFALL`, `TREEMAP`, `SUNBURST`,
`HISTOGRAM`, `PARETO_LINE`) live in their own `ChartType` values.

## Quick example — pick based on data shape

```java
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().get(0);

int topRow = 0, leftCol = 0, bottomRow = 15, rightCol = 8;
boolean categorical  = true;   // from inspecting your data
boolean timeSeries   = false;
boolean shareOfTotal = false;
boolean comparative  = false;

int type = categorical  ? ChartType.COLUMN
         : timeSeries   ? ChartType.LINE
         : shareOfTotal ? ChartType.PIE
         : comparative  ? ChartType.BAR_STACKED
                        : ChartType.COLUMN;

int idx = ws.getCharts().add(type, topRow, leftCol, bottomRow, rightCol);
Chart chart = ws.getCharts().get(idx);
```

`ChartType` is a class of `int` constants — assign to `int`, not to a
`ChartType` variable.

## Combo chart

```java
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().get(0);
int comboIdx = ws.getCharts().add(ChartType.COLUMN, 0, 0, 20, 10);
Chart combo = ws.getCharts().get(comboIdx);
combo.getNSeries().add("Sheet1!$B$1:$B$10", true);   // bars
combo.getNSeries().add("Sheet1!$C$1:$C$10", true);   // lines on top
combo.getNSeries().get(1).setPlotOnSecondAxis(true);
```

`setPlotOnSecondAxis(true)` moves the series onto the chart's
secondary value axis.

## Pitfalls

- **`ChartType.COLUMN` vs `ChartType.BAR`** — `COLUMN` is vertical
  bars (like Excel's "Column"); `BAR` is horizontal (Excel's "Bar").
  Many users have it reversed.
- **3D variants** — in Excel 2016+ dropdowns 3D is sometimes hidden
  under their sub-menu; here it is a separate enum value.
- **`_STACKED` vs `_100_PERCENT_STACKED`** — `_STACKED` (e.g.
  `BAR_STACKED`) sums absolute; `_100_PERCENT_STACKED` normalises so
  each bar always adds to 100 %.
- **Funnel / Waterfall / Box & Whisker** — only render correctly
  with one column of data plus optional category labels.

## Related

- [create-chart.md](create-chart.md)
- [chart-data-formatting.md](chart-data-formatting.md)
