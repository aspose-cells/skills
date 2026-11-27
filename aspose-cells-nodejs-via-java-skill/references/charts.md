# Charts

## Create and populate a chart

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook();
const sheet = workbook.getWorksheets().get(0);
const data = [
  ["Region", "Revenue"],
  ["North", 125],
  ["South", 99],
  ["West", 140]
];
for (let r = 0; r < data.length; r += 1) {
  sheet.getCells().get(r, 0).putValue(data[r][0]);
  sheet.getCells().get(r, 1).putValue(data[r][1]);
}

const index = sheet.getCharts().add(cells.ChartType.COLUMN, 5, 0, 20, 8);
const chart = sheet.getCharts().get(index);
chart.getNSeries().add("B2:B4", true);
chart.getNSeries().setCategoryData("A2:A4");
chart.getTitle().setText("Revenue");
workbook.calculateFormula();
chart.calculate();
workbook.save("chart.xlsx", cells.SaveFormat.XLSX);
```

Chart anchors are 0-based worksheet positions. Add series before configuring series-specific formatting. Qualify ranges with a quoted sheet name when the source is on another worksheet.

## Render a chart

Calculate after changing source values, formulas, series, or categories. Otherwise chart images can be empty or stale.

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook("chart.xlsx");
const chart = workbook.getWorksheets().get(0).getCharts().get(0);
workbook.calculateFormula();
chart.calculate();
chart.toImage("chart.png");
chart.toPdf("chart.pdf");
```

Use version-specific `ImageOrPrintOptions` when output dimensions, format, or quality must be controlled. Install source fonts in the rendering environment and test text clipping on the production OS.

## Troubleshooting

- Blank chart: verify source ranges, cached values, and the `calculate()` call.
- Wrong categories: set category data after adding all intended series and confirm the range orientation.
- Stale chart: recalculate formula precedents before calculating the chart.
- Wrong labels or layout: check fonts and dimensions rather than adding arbitrary padding.
- Saving a workbook is not the same operation as rendering a chart; test both outputs when both are required.
