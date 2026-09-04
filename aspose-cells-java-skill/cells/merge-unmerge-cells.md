---
name: merge-unmerge-cells
description: Merge and unmerge cell ranges in a worksheet with Aspose.Cells for Java, including merged-cell validation and named-range preservation. Use when the user wants to combine cells for headers or layouts.
applies_to: Java
keywords: [merge, unmerge, range, header]
---

# Merge & unmerge cells

## Imports

```java
import com.aspose.cells.*;
```

## Quick example

```java
Worksheet ws = wb.getWorksheets().get(0);

// Merge A1:E1
ws.getCells().merge(0, 0, 0, 4);

// Unmerge later
ws.getCells().unMerge(0, 0, 0, 4);
```

`merge(startRow, startColumn, endRow, endColumn)` — **all zero-based**.

If you need to merge only when no data would be lost, use
`merge(startRow, ..., mergeConflict = true)` (default) — passing
`false` aborts the merge when the range isn't empty.

## Range API

```java
Range title = ws.getCells().createRange("A1:E1");
title.merge();                                     // merge via Range
```

## Read merged cells

```java
CellArea[] merged = ws.getCells().getMergedAreas();
for (CellArea area : merged) {
    int r1 = area.StartRow, r2 = area.EndRow, c1 = area.StartColumn, c2 = area.EndColumn;
}
```

## Pitfalls

- **Merged cells keep only the top-left value** — values in the
  lower-right cells are silently discarded. Save first if that
  matters.
- **Defined names / data validators** pointing into a merged range
  may behave oddly — recheck after merge/unmerge.
- **HTML/PDF render respects merges** — but auto-fit widths ignore
  merged cells unless the width was set explicitly.
- **Tests for "everything merged"** — use `getMergedAreas()`,
  *not* iterating rows.

## Related

- [read-write-values.md](read-write-values.md)
- [page-setup-print.md](../worksheet/page-setup-print.md)
