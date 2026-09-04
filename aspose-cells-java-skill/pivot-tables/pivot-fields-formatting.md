---
name: pivot-fields-formatting
description: Configure pivot field buckets — display subtotals / grand totals, value-field functions (SUM/AVG/COUNT), sort, filter, top-N, drill-down, conditional formats. Use whenever the user wants to fine-tune an existing pivot or change aggregation behaviour.
applies_to: Java
keywords: [pivot field, subtotal, grand total, value field, consolidation function, sort, filter, topN]
---

# Pivot field configuration

## Imports

```java
import com.aspose.cells.*;
```

## Add fields to areas

```java
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().add("Summary");
int pIdx = ws.getPivotTables().add("=Sheet1!$A$1:$C$5", "A3", "MyPivot");
PivotTable pt = ws.getPivotTables().get(pIdx);

pt.addFieldToArea(PivotFieldType.PAGE,   1);    // report filter
pt.addFieldToArea(PivotFieldType.ROW,    0);    // rows
pt.addFieldToArea(PivotFieldType.COLUMN, 2);    // columns
pt.addFieldToArea(PivotFieldType.DATA,   3);    // values
```

## Value field functions

```java
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().add("Summary");
int pIdx = ws.getPivotTables().add("=Sheet1!$A$1:$C$5", "A3", "MyPivot");
PivotTable pt = ws.getPivotTables().get(pIdx);
pt.addFieldToArea(PivotFieldType.DATA, 2);

PivotField data = pt.getDataFields().get(0);
data.setFunction(ConsolidationFunction.SUM);
data.setFunction(ConsolidationFunction.AVERAGE);
data.setFunction(ConsolidationFunction.COUNT);
data.setFunction(ConsolidationFunction.MAX);
data.setFunction(ConsolidationFunction.MIN);
data.setFunction(ConsolidationFunction.PRODUCT);
data.setFunction(ConsolidationFunction.STD_DEV);
// ... STD_DEVP, VAR, VARP, DISTINCT_COUNT all in the enum.

data.setName("Total Amount");
data.setNumberFormat("#,##0.00");
```

## Subtotals & grand totals

```java
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().add("Summary");
int pIdx = ws.getPivotTables().add("=Sheet1!$A$1:$C$5", "A3", "MyPivot");
PivotTable pt = ws.getPivotTables().get(pIdx);
pt.addFieldToArea(PivotFieldType.ROW, 0);

pt.setShowRowGrandTotals(true);
pt.setShowColumnGrandTotals(true);

PivotField region = pt.getRowFields().get(0);
region.setSubtotals(PivotFieldSubtotalType.SUM, true);
region.setSubtotals(PivotFieldSubtotalType.COUNT, false);
```

## Sort / filter a field

```java
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().add("Summary");
int pIdx = ws.getPivotTables().add("=Sheet1!$A$1:$C$5", "A3", "MyPivot");
PivotTable pt = ws.getPivotTables().get(pIdx);
pt.addFieldToArea(PivotFieldType.ROW, 0);

PivotField year = pt.getRowFields().get(0);
year.setAutoSortField(0);                                 // by the field itself
year.setAscendSort(true);

year.getPivotItems().get(0).setHidden(true);             // hide the 1st item (e.g., 2019)
year.getPivotItems().get(1).setHidden(true);             // hide the 2nd item (e.g., 2020)
```

`field.getPivotItems().get(i).setHidden(hidden)` hides the item at
index `i` in the field. For text labels, the `String` overload
(`field.hideItem(String itemValue, boolean isHidden)`) is still the
shortest form.

## Top-N

```java
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().add("Summary");
int pIdx = ws.getPivotTables().add("=Sheet1!$A$1:$C$5", "A3", "MyPivot");
PivotTable pt = ws.getPivotTables().get(pIdx);
pt.addFieldToArea(PivotFieldType.ROW, 0);

PivotField field = pt.getRowFields().get(0);
field.setAutoShow(true);                // enable top-N filter
field.setAutoShowCount(10);              // show the top 10
field.setAutoShowField(0);              // rank by this field
```

## Conditional formatting pivot values

```java
import static com.aspose.cells.Color.fromArgb;
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().add("Summary");
int pIdx = ws.getPivotTables().add("=Sheet1!$A$1:$C$5", "A3", "MyPivot");
PivotTable pt = ws.getPivotTables().get(pIdx);

ConditionalFormattingCollection cfc = ws.getConditionalFormattings();
int cIdx = cfc.add();
FormatConditionCollection fc = cfc.get(cIdx);

CellArea area = pt.getDataBodyRange();
fc.addArea(area);
fc.addCondition(FormatConditionType.COLOR_SCALE);
fc.get(0).getColorScale().setMaxColor(fromArgb(0xC00000));
```

The `ColorScale` accessor comes back on the `FormatCondition` directly —
no cast needed. The `import static` disambiguates `Color` from the
prelude's `java.awt.Color`.

## Pitfalls

- **A field placed in `DATA` becomes a numeric field by default**;
  text columns get `Count` automatically. Set the function explicitly
  to avoid confusion.
- **`setAutoSortField` index is 0-based within the data area** —
  different from the source range.
- **`setSubtotals` is additive** — pass each subtotal type
  individually; passing one type leaves others untouched.
- **`addFieldToArea` over a `DATA` field changes the consolidation
  function default** — re-apply after the call.

## Related

- [create-pivot-table.md](create-pivot-table.md)
- [refresh-pivot-table.md](refresh-pivot-table.md)
