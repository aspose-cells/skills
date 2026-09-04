---
name: refresh-pivot-table
description: Refresh pivot tables in Aspose.Cells for Java — workbook.refreshAll / worksheet.refreshPivotTables / pivotTable.refreshData / pivotTable.calculateData, layered refresh strategy, deferred cache, and how to tell which method to use.
applies_to: Java
keywords: [refresh, refreshAll, refreshPivotTables, pivot cache, refreshData, deprecated]
---

# Refresh a pivot table

`pivotTable.getPivotCache().refresh()` re-reads the source range every
call; for routine recomputes after in-memory edits prefer
`calculateData()`. Use the layered refresh API for the appropriate
scope:

| Need                                     | Method                            |
| ---------------------------------------- | --------------------------------- |
| Refresh everything (workbook)            | `workbook.refreshAll()`           |
| Refresh all pivots on a sheet            | `worksheet.refreshPivotTables()`  |
| Refresh one pivot's cache                | `pivotTable.getPivotCache().refresh()` |
| Recompute aggregates from current cache  | `pivotTable.calculateData()`      |
| Rebuild cache + recompute in one call    | `pivotTable.calculateData(opt)` where `opt.setRefreshData(true)` |

## Imports

```java
import com.aspose.cells.*;
```

## Quick example

```java
Workbook wb = new Workbook("data.xlsx");
// ... edit source data

wb.refreshAll();                             // fastest for the whole file
wb.save("with-fresh-pivots.xlsx");
```

## Sheet-scoped refresh

```java
Worksheet ws = wb.getWorksheets().get("Summary");
ws.refreshPivotTables();
```

## Just one pivot, fast path

```java
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().add("Summary");
int pIdx = ws.getPivotTables().add("=Sheet1!$A$1:$C$5", "A3", "MyPivot");
PivotTable pt = ws.getPivotTables().get(pIdx);

pt.getPivotCache().refresh();                // recomputes aggregates; faster
                                             // than reloading the source
```

Use `pt.getPivotCache().refresh()` only after you replaced the
source range or changed its contents from outside the workbook
(e.g. an SQL query).

## Pitfalls

- **Use `pt.getPivotCache().refresh()` only for source swaps.** Routine edits
  followed by `refresh()` re-reads the entire source — slow.
- **Cached pivot source range** — `pt.getSource()` may be a
  different snapshot than the current `changeDataSource(...)` value;
  update it if the source range moved.
- **External (DB / OLAP) sources** cannot be refreshed by
  `pivotTable.getPivotCache().refresh()`. Use the connection's own refresh.
- **Refresh after editing values, before saving.** Otherwise the
  saved file ships stale aggregates.

### Body is empty / shows stale numbers after source edits

Symptom: after `pivotTable.add(...)` or after editing the source
cells, the pivot area is blank, shows only field headers, or keeps
the numbers from a previous run.

Cause: the pivot body is computed on demand from the pivot cache.
`add(...)` builds the cache once from the source at that moment —
later edits to the source cells do not flow into the cache, and
Aspose never recomputes on save.

Fix, in order:

1. If the source cells are formula-driven, call
   `wb.calculateFormula()` first so the source values are current.
2. Rebuild the cache from the source AND recompute the view in one
   call: `pt.calculateData(new PivotTableCalculateOption())` with
   `setRefreshData(true)`. Equivalent to the longer
   `pt.getPivotCache().refresh(); pt.calculateData();` pair, but in
   one call.
3. Do this before `save()` and before reading any pivot output cell.

```java
Workbook wb = new Workbook("report.xlsx");
PivotTable pt = wb.getWorksheets().get(0).getPivotTables().get(0);

// 1. source values are current (skip if source is plain data)
wb.calculateFormula();

// 2. rebuild cache from current source + recompute, in one call
PivotTableCalculateOption opt = new PivotTableCalculateOption();
opt.setRefreshData(true);
pt.calculateData(opt);                       // returns PivotTable[] (dependents)

wb.save("report.xlsx");
```

`pt.calculateData()` with no option recomputes only from the
**existing** cache. That is fine immediately after `add(...)` (the
cache was just built), but it will NOT pick up later edits to the
source cells — those need `setRefreshData(true)`.

### Pivot body cells read as null before calculation

Symptom: reading cells in the pivot area returns `null` / empty
even though the pivot looks set up.

Cause: pivot output cells are ordinary cells only *after* the body
is computed.

Fix: call `calculateData(...)` (see above), then read via the
pivot's ranges. `DataBodyRange` is the values region;
`TableRange1` is the whole report excluding page fields;
`TableRange2` includes page fields. All expose `CellArea` with
`startRow` / `startColumn` / `endRow` / `endColumn`.

```java
Workbook wb = new Workbook("report.xlsx");
PivotTable pt = wb.getWorksheets().get(0).getPivotTables().get(0);
PivotTableCalculateOption opt = new PivotTableCalculateOption();
opt.setRefreshData(true);
pt.calculateData(opt);                       // body cells are null before this

com.aspose.cells.CellArea body = pt.getDataBodyRange();   // values region
Object v = wb.getWorksheets().get(0).getCells().get(body.StartRow, body.StartColumn).getValue();
```

### Defer refresh to Excel on open

Symptom: the pivot depends on a source you cannot fully reproduce
headless, or you want the numbers to be live when the user opens
the file.

Fix: set `pt.setRefreshDataOnOpeningFile(true)`. Excel then
refreshes the pivot when the file is opened. Note this does not
populate the body inside the saved file — a headless reader
(including Aspose re-opening it) still sees the old cache until it
runs `calculateData(...)`.

## Related

- [create-pivot-table.md](create-pivot-table.md)
- [pivot-fields-formatting.md](pivot-fields-formatting.md)
