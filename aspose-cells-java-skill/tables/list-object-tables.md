---
name: list-object-tables
description: Add Excel-style ListObject tables in Aspose.Cells for Java — CreateListObject, table style, totals row, calculated columns, structured references in formulas.
applies_to: Java
keywords: [list object, table, CreateListObject, structured reference, totals row]
---

# ListObject tables (Excel Tables)

## Imports

```java
import com.aspose.cells.*;
```

## Quick example

```java
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().get(0);

// Data with header
ws.getCells().importTwoDimensionArray(new String[][] {
    {"Item",  "Qty", "Price"},
    {"Apple", "10",  "1.20"},
    {"Pear",  "5",   "0.80"},
    {"Grape", "30",  "0.50"},
}, 0, 0, false);

int idx = ws.getListObjects().add(0, 0, 3, 2, true);   // startRow, startCol, endRow, endCol, hasHeaders
ListObject table = ws.getListObjects().get(idx);

table.setShowTotals(true);
table.getListColumns().get(0).setTotalsRowLabel("Total");
table.getListColumns().get(1).setCustomTotalsRowFormula("=SUBTOTAL(109,Inventory[Qty])", true, true);
table.getListColumns().get(2).setCustomTotalsRowFormula("=SUBTOTAL(109,Inventory[Price])", true, true);
table.setTableStyleName("TableStyleMedium2");

wb.save("table.xlsx");
```

## Calculated column

```java
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().get(0);
int idx = ws.getListObjects().add(0, 0, 3, 2, true);   // include header row
ListObject table = ws.getListObjects().get(idx);
table.getListColumns().get(2).setFormula("=Inventory[Price]*0.2");   // Tax column
```

The header row is taken from row 0. Calculated columns extend
automatically when users enter data into new rows.

## Structured references

```java
ws.getCells().get("E2").setFormula("=SUM(Inventory[Qty])");
```

The table name `Inventory` becomes a global range; columns within it
are addressed by name (`[Qty]`).

## Pitfalls

- **`hasHeaders` must match the actual data**. Lying (`true` when
  there is no header) causes the first row to become the column name
  — usually `<empty>`.
- **Structured references** are case-sensitive in formulas — use the
  same casing as the column header text.
- **Adding a totals row** when `setShowTotals(false)` silently keeps
  the previous state; the total formula requires `setShowTotals(true)`.
- **Tables cannot span multiple sheets** — one `ListObject` per
  worksheet source range.

## Related

- [../cells/cell-formatting-styles.md](../cells/cell-formatting-styles.md)
- [../cells/read-write-values.md](../cells/read-write-values.md)
