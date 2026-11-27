# Pivot tables

## Create a pivot from worksheet data

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook();
const dataSheet = workbook.getWorksheets().get(0);
dataSheet.setName("Data");
const rows = [
  ["Region", "Amount"],
  ["North", 125],
  ["South", 99],
  ["North", 75]
];
for (let r = 0; r < rows.length; r += 1) {
  dataSheet.getCells().get(r, 0).putValue(rows[r][0]);
  dataSheet.getCells().get(r, 1).putValue(rows[r][1]);
}

const outputIndex = workbook.getWorksheets().add();
const output = workbook.getWorksheets().get(outputIndex);
output.setName("Summary");
const index = output.getPivotTables().add("Data!A1:B4", "A3", "SalesPivot");
const pivot = output.getPivotTables().get(index);
pivot.addFieldToArea(cells.PivotFieldType.ROW, 0);
pivot.addFieldToArea(cells.PivotFieldType.DATA, 1);
pivot.getPivotCache().refresh();
pivot.calculateData();
workbook.save("pivot.xlsx", cells.SaveFormat.XLSX);
```

The source range must include headers. Use 0-based source-field indices in `addFieldToArea()` and quote source sheet names containing spaces.

## Refresh existing pivots

Changing source cells does not automatically rebuild the pivot cache or output.

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook("pivot-input.xlsx");
const pivot = workbook.getWorksheets().get("Summary").getPivotTables().get(0);
pivot.getPivotCache().refresh();
pivot.calculateData();
workbook.save("pivot-output.xlsx", cells.SaveFormat.XLSX);
```

Refresh the cache before calculating data. If the source range itself grew, update the cache source as required by the installed API before refreshing; refresh alone does not infer new source bounds.

## Common failures

- Empty output: no fields were added, the cache was not refreshed, or `calculateData()` was omitted.
- Stale totals: source values changed without a cache refresh and recalculation.
- Wrong aggregation: inspect each data field's consolidation function instead of assuming every field defaults to sum.
- Missing new rows: expand the source range or use an appropriate source table before refreshing.
- Multiple pivots: determine whether caches are shared before changing a cache source or refresh policy.
