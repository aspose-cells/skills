---
name: sort-data
description: Sort cell ranges in Aspose.Cells for Java by one or more keys, ascending / descending, with custom sort lists or header-row exclusion. Use whenever the user asks to sort rows of data.
applies_to: Java
keywords: [sort, order, ascending, descending, custom list]
---

# Sort data

## Imports

```java
import com.aspose.cells.*;
```

## Quick example — one key

```java
Worksheet ws = wb.getWorksheets().get(0);
DataSorter sorter = wb.getDataSorter();

sorter.setOrder1(SortOrder.ASCENDING);
sorter.setKey1(0);                  // first column in the range

CellArea area = new CellArea();
area.StartRow = 1; area.EndRow = 100;
area.StartColumn = 0; area.EndColumn = 4;
sorter.sort(ws.getCells(), area);
```

## Sort with a header row

```java
DataSorter sorter = wb.getDataSorter();
sorter.setHasHeaders(true);                     // row 1 is the header
CellArea area = new CellArea();
area.StartRow = 1; area.EndRow = 100;
area.StartColumn = 0; area.EndColumn = 4;
sorter.sort(ws.getCells(), area);
```

## Multi-key + custom sort list

```java
DataSorter sorter = wb.getDataSorter();

// Two sort keys: col A (descending), col C (ascending)
sorter.addKey(0, SortOrder.DESCENDING);
sorter.addKey(2, SortOrder.ASCENDING);

// Custom: strings listed in this order come first (per key)
String[] custom = {"Critical", "High", "Normal", "Low"};
sorter.addKey(0, SortOrder.ASCENDING, custom);

CellArea area = new CellArea();
area.StartRow = 1; area.EndRow = 100;
area.StartColumn = 0; area.EndColumn = 4;
sorter.sort(ws.getCells(), area);
```

## Pitfalls

- **`setHasHeaders(true)` excludes the first row from sort** — for
  an unsorted file that has no headers, leave it off.
- **Keys are range-relative 0-based indexes**, not column letters.
- **Cell formulas referencing sorted rows are not rewritten** —
  external formulas break. Sort within a sheet only.
- **`DataSorter` is a workbook-level instance** — reuse it for many
  sheets.
- **Warning callbacks** — register an `IWarningCallback` on the
  workbook via `wb.getSettings().setWarningCallback(...)` if the
  range contains formulas or merged cells.

## Related

- [auto-filter.md](auto-filter.md)
- [find-search-replace.md](find-search-replace.md)
