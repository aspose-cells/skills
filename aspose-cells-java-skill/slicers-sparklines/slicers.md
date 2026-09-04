---
name: slicers
description: Add interactive slicers (visual filters) over Excel tables (ListObject) and pivot tables with Aspose.Cells for Java. Use when the user wants a click-to-filter dashboard control on a sheet, or to read/write the slicer cache items.
applies_to: Java
keywords: [slicer, SlicerCollection, SlicerCache, ListObject, pivot table, filter, dashboard]
---

# Slicers

Slicers are visual filter buttons that attach to an Excel table
(`ListObject`) or a pivot table (`PivotTable`). They live on
`Worksheet.getSlicers()` (a `SlicerCollection`).

## Imports

```java
import com.aspose.cells.*;
```

## Attach a slicer to an Excel table (`ListObject`)

`SlicerCollection.add(ListObject, int columnIndex, String destCellName)`
returns an int index; index the collection to configure the `Slicer`.
The destination cell is A1-style (`"E2"`).

```java
Workbook wb = new Workbook("input.xlsx");
Worksheet ws = wb.getWorksheets().get(0);

// Assumes a ListObject already exists on the sheet — see list-object-tables.md.
ListObject table = ws.getListObjects().get(0);

// Slicer over the table's column 0 (0-based), anchored at E2.
int idx = ws.getSlicers().add(table, 0, "E2");
Slicer slicer = ws.getSlicers().get(idx);

slicer.setStyleType(SlicerStyleType.SLICER_STYLE_LIGHT_1);
slicer.setCaption("Region");
slicer.setNumberOfColumns(2);   // lay out filter buttons in 2 columns
slicer.setColumnWidth(90);      // px per button

wb.save("out.xlsx");
```

The available styles are constants on `SlicerStyleType`
(`SLICER_STYLE_LIGHT_1` … `SLICER_STYLE_LIGHT_6`,
`SLICER_STYLE_DARK_1` … `SLICER_STYLE_DARK_6`,
`SLICER_STYLE_OTHER_1`, `SLICER_STYLE_OTHER_2`, `CUSTOM`).

## Attach a slicer to a pivot table

Requires the pivot to exist first (see
[../pivot-tables/create-pivot-table.md](../pivot-tables/create-pivot-table.md)).
Bind to a pivot field by 0-based index or by name.

```java
Workbook wb = new Workbook("report.xlsx");
Worksheet ws = wb.getWorksheets().get(0);
PivotTable pt = ws.getPivotTables().get(0);

// Bind to the pivot field at 0-based index 0; anchor at F2.
int idx = ws.getSlicers().add(pt, 0, 2, "F2");
Slicer slicer = ws.getSlicers().get(idx);
slicer.setStyleType(SlicerStyleType.SLICER_STYLE_DARK_1);
wb.save("out.xlsx");
```

There are several pivot-bound overloads:

| Overload                                              | Use                                              |
| ----------------------------------------------------- | ------------------------------------------------ |
| `add(PivotTable, int fieldIndex, int row, int col)`   | bind by field 0-based index                      |
| `add(PivotTable, String fieldName, int row, int col)` | bind by field name                               |
| `add(PivotTable, int fieldIndex, PivotField)`         | reuse an existing `PivotField`                   |
| `add(PivotTable, String fieldName, PivotField)`       | reuse an existing `PivotField` by name           |
| `add(PivotTable, String destCell, String fieldName)`  | A1-style anchor                                  |
| `add(PivotTable, String destCell, int fieldIndex)`    | A1-style anchor, 0-based field                   |

## Reading and writing selected items

A slicer's allowed values live in `Slicer.getSlicerCache()` (a
`SlicerCache`), with `getSlicerCacheItems()` returning a
`SlicerCacheItemCollection`. Each item has `getValue()` (the filter
value) and `getSelected()` (whether it is currently selected). The
cache is shared between a slicer and its bound pivot field, so
toggling it here affects the pivot's filter.

```java
Workbook wb = new Workbook("report.xlsx");
Worksheet ws = wb.getWorksheets().get(0);
Slicer slicer = ws.getSlicers().get(0);

SlicerCacheItemCollection items = slicer.getSlicerCache().getSlicerCacheItems();
for (int i = 0; i < items.getCount(); i++) {
    SlicerCacheItem item = items.get(i);
    System.out.println(item.getValue() + " selected=" + item.getSelected());
}

// Toggle programmatically — also set/unset by item value via selectItems.
slicer.selectItems(new String[]{"East"}, true);    // restrict to "East"
slicer.unselectItems(new String[]{"West"});
slicer.clearFilter();                              // reset to all selected
```

## Pitfalls

### Slicer over a table needs the table, not raw cells

Symptom: `ws.getSlicers().add(...)` throws, or the slicer shows nothing
after save.
Cause: slicers attach to a `ListObject` or a `PivotTable`, not to an
arbitrary cell range. A range that is not a real Excel table has no
slicer model.
Fix: create a `ListObject` first via `ws.getListObjects().add(...)`
(see [../tables/list-object-tables.md](../tables/list-object-tables.md)),
or bind to a pivot instead.

### Slicer over a pivot does not filter the sheet until the pivot is recalculated

Symptom: the slicer exists but filtering does nothing to the rows /
totals you read.
Cause: a slicer over a pivot only filters the pivot cache; the pivot
body (and any formulas over it) must be recalculated after the
selection changes. Aspose.Cells does not do this for you.
Fix: after toggling `SlicerCacheItem.setSelected(false)`, refresh the
pivot data and recalculate the workbook, then re-read the affected
cells.

```java
Workbook wb = new Workbook("report.xlsx");
Worksheet ws = wb.getWorksheets().get(0);
PivotTable pt = ws.getPivotTables().get(0);
Slicer slicer = ws.getSlicers().get(0);

slicer.getSlicerCache().getSlicerCacheItems().get(0).setSelected(false);

PivotTableCalculateOption opt = new PivotTableCalculateOption();
opt.setRefreshData(true);
pt.calculateData(opt);
wb.calculateFormula();
// ... now read the filtered pivot cells
```

## Removing slicers

`SlicerCollection` exposes `remove(Slicer)`, `removeAt(int)`, and
`clear()`. Use `remove` for by-reference removal, `removeAt` for
positional.

```java
Workbook wb = new Workbook("with-slicer.xlsx");
Worksheet ws = wb.getWorksheets().get(0);

if (ws.getSlicers().getCount() > 0) {
    ws.getSlicers().removeAt(0);
}
wb.save("out.xlsx");
```

## Notes

- A slicer is **not** a chart — it does not need `chart.calculate()`
  before rendering. The shape is rendered by Excel directly.
- `Slicer.refresh()` rebuilds the slicer cache from the bound source.
  Call it after editing source data on the same sheet if the slicer is
  attached to a `ListObject`.

## Related

- [../pivot-tables/create-pivot-table.md](../pivot-tables/create-pivot-table.md) —
  the pivot a slicer filters; refresh before reading filtered output.
- [../tables/list-object-tables.md](../tables/list-object-tables.md) —
  `ListObject` tables that slicers attach to.
- [sparklines.md](sparklines.md) — mini in-cell charts.
