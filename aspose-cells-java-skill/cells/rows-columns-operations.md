---
name: rows-columns-operations
description: Insert, delete, copy, hide, group, and resize rows and columns in Aspose.Cells for Java. Use whenever the user asks to add or remove rows / columns, freeze panes, group rows, or adjust widths / heights.
applies_to: Java
keywords: [rows, columns, insert, delete, hide, group, height, width]
---

# Rows & columns operations

## Imports

```java
import com.aspose.cells.*;
```

## Quick example

```java
Worksheet ws = wb.getWorksheets().get(0);
Cells cells = ws.getCells();

// Insert 5 empty rows at index 2
cells.insertRows(2, 5);

// Delete 3 columns starting at column 5 (E)
cells.deleteColumns(4, 3, true);

// Copy row 1..3 above row 10 (keeping formulas/styles)
cells.copyRows(cells, 0, 9, 3);
```

## Column / row width and visibility

```java
Worksheet ws = wb.getWorksheets().get(0);
Cells cells = ws.getCells();
cells.setColumnWidth(0, 25);                       // column A
cells.setRowHeight(0, 30);

cells.hideColumn(2);
cells.hideRow(5);
cells.unhideColumn(2, 25.0);                      // (colIndex, width)
```

## Group / outline

```java
Worksheet ws = wb.getWorksheets().get(0);
Cells cells = ws.getCells();
cells.groupRows(5, 20, true);                      // group rows 6..21, summary below
cells.groupColumns(2, 4, true);                    // group C..E, summary right
cells.getRows().get(5).setCollapsed(true);         // collapse that group
```

## Insert / delete cells (not whole rows)

```java
Worksheet ws = wb.getWorksheets().get(0);
Cells cells = ws.getCells();
// Move A1's content one cell right and shift neighbours
CellArea area = new CellArea();
area.StartRow = 0; area.EndRow = 0;
area.StartColumn = 0; area.EndColumn = 0;
cells.insertRange(area, 0, 1, true);               // rowOff=0, colOff=1, updateReference
cells.deleteRange(0, 0, 0, 0, 0);                  // shiftType 0 = UP
```

## Pitfalls

- **`insertRows(idx, 1)` updates `cell` references** in the
  worksheet but **does not** update `chart` data sources or named
  ranges — rebuild them or call `wb.refreshAll()` afterwards if
  relevant.
- **`deleteRows` shift formulas down** — absolute references inside
  the deleted range become `#REF!`.
- **Grouping uses 1-based row indexes in formulas** — be careful when
  mixing programmatic indexes.
- **Hidden rows still iterate** — use
  `cells.getRows().get(5).isHidden()` to filter, or pass
  `new PasteOptions().setOnlyVisibleCells(true)` to skip hidden rows
  when pasting.
- **Column widths are in characters**, not pixels. Multiply by ~7 for
  rough pixel estimate.

## Related

- [merge-unmerge-cells.md](merge-unmerge-cells.md)
- [panes-zoom-views.md](../worksheet/panes-zoom-views.md)
