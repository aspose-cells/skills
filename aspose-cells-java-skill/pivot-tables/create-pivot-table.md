---
name: create-pivot-table
description: Create a pivot table in Aspose.Cells for Java from a worksheet data range — choose fields, row/column/data areas, build the pivot. Use when the user wants any kind of pivot table or PivotTable/CrossTab on a sheet.
applies_to: Java
keywords: [pivot table, addFieldToArea, pivotFields, cross-tab, data area]
---

# Create a pivot table

## Imports

```java
import com.aspose.cells.*;
```

## Quick example

```java
Workbook wb = new Workbook();
Worksheet source = wb.getWorksheets().get(0);

// 1) Data
source.getCells().get("A1").setValue("Region");   source.getCells().get("B1").setValue("Product");
source.getCells().get("A2").setValue("East");     source.getCells().get("B2").setValue("Apples");
source.getCells().get("A3").setValue("West");     source.getCells().get("B3").setValue("Pears");
source.getCells().get("A4").setValue("East");     source.getCells().get("B4").setValue("Apples");
source.getCells().get("A5").setValue("West");     source.getCells().get("B5").setValue("Pears");
source.getCells().get("C1").setValue("Amount");
source.getCells().get("C2").setValue(1500);
source.getCells().get("C3").setValue(2400);
source.getCells().get("C4").setValue(1800);
source.getCells().get("C5").setValue(2200);

// 2) Pivot on a *different* sheet (recommended)
Worksheet piv = wb.getWorksheets().add("Summary");
int idx = piv.getPivotTables().add("=Sheet1!$A$1:$C$5", "A3", "MyPivot");
PivotTable pt = piv.getPivotTables().get(idx);

// 3) Layout
pt.addFieldToArea(PivotFieldType.ROW, 0);     // Region
pt.addFieldToArea(PivotFieldType.COLUMN, 1);  // Product
pt.addFieldToArea(PivotFieldType.DATA, 2);    // Amount (sum)

pt.setShowRowGrandTotals(true);
pt.setShowColumnGrandTotals(true);

wb.save("pivot.xlsx");
```

## Choose fields programmatically (post-creation)

```java
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().add("Summary");
int pIdx = ws.getPivotTables().add("=Sheet1!$A$1:$C$5", "A3", "MyPivot");
PivotTable pt = ws.getPivotTables().get(pIdx);

PivotFieldCollection fields = pt.getRowFields();
fields.addByBaseIndex(0);                  // first source column
fields.addByBaseIndex(2);                  // third source column

PivotField valueField = pt.getDataFields().get(0);
valueField.setFunction(ConsolidationFunction.SUM);
valueField.setNumberFormat("0.00");
```

## Make a pivot chart too

```java
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().add("Summary");
int pIdx = ws.getPivotTables().add("=Sheet1!$A$1:$C$5", "A3", "MyPivot");
PivotTable pt = ws.getPivotTables().get(pIdx);

// Anchor the chart to the right of the pivot body (use row 2, col E)
int chartIdx = ws.getCharts().add(ChartType.BAR, 2, 5, 22, 20);
ws.getCharts().get(chartIdx).setPivotSource("MyPivot");
```

For programmatic positioning, use `PivotTable.getColumnRange()` /
`getRowRange()` (return `CellArea`) to derive offsets from the pivot
body.

## Pitfalls

- **Source must be top-left → bottom-right**. `"C5:A1"` throws.
- **First call to `addFieldToArea` registers the field, subsequent
  calls for the same field are no-ops** — use `pt.setDataFields`
  only when overriding the auto-AVERAGE default.
- **Source cell range must include the header row** — otherwise
  column A would appear named like "Region" in the field list (not
  the actual column label).
- **Don't put the pivot on the same sheet as the data unless the user
  explicitly asks** — Excel disallows a pivot overlapping its source
  range, but Aspose writes it; opening in Excel fails.
- **Refresh cache after changing source values** — see
  `refresh-pivot-table.md`.

## Related

- [pivot-fields-formatting.md](pivot-fields-formatting.md)
- [refresh-pivot-table.md](refresh-pivot-table.md)
