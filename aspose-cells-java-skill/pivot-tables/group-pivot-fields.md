---
name: group-pivot-fields
description: Group pivot fields in Aspose.Cells for Java — dates by years/quarters/months, numbers into fixed ranges, auto-grouping, ungroup. Use when a pivot should aggregate by date period or numeric bin instead of raw values.
applies_to: Java
keywords: [pivot group, group field, PivotFieldGroupOptions, PivotGroupByType, date grouping, numeric range, ungroup]
---

# Group pivot fields

Grouping happens on the `PivotField`, not on the `PivotTable`. Build a
`PivotFieldGroupOptions`, set the ranges / group types, then call
`field.group(options)` — or `field.group()` to let Aspose pick the
groups. Recompute with `pt.calculateData()` (or `pt.refresh()`) and
save.

## Imports

```java
import com.aspose.cells.*;
```

## Group a date field (years / quarters / months)

```java
Workbook wb = new Workbook("report.xlsx");
PivotTable pt = wb.getWorksheets().get(0).getPivotTables().get(0);

PivotField dateField = pt.getRowFields().get(0);
PivotFieldGroupOptions grp = new PivotFieldGroupOptions();
grp.setGroupTypes(new int[]{
        PivotGroupByType.YEARS, PivotGroupByType.QUARTERS, PivotGroupByType.MONTHS});
grp.setStartDate(new DateTime(2026, 1, 1));       // optional bounds
grp.setEndDate(new DateTime(2026, 12, 31));
dateField.group(grp);

pt.calculateData();
wb.save("grouped.xlsx");
```

Omit `setStartDate` / `setEndDate` to derive the range from the data.

## Auto-group

```java
PivotTable pt = ws.getPivotTables().get(0);
PivotField f = pt.getRowFields().get(0);
f.group();                    // Aspose chooses date parts or numeric bins
pt.calculateData();
```

## Numeric ranges

```java
PivotTable pt = ws.getPivotTables().get(0);
PivotField amount = pt.getRowFields().get(0);
PivotFieldGroupOptions bins = new PivotFieldGroupOptions();
bins.setGroupTypes(new int[]{PivotGroupByType.NUMBERS});
bins.setStartNumber(0.0);
bins.setEndNumber(1000.0);
bins.setInterval(100.0);      // buckets: 0-100, 100-200, ...
amount.group(bins);
pt.calculateData();
```

## Ungroup

```java
PivotTable pt = ws.getPivotTables().get(0);
pt.getRowFields().get(0).ungroup();
pt.calculateData();
```

Read back the live grouping with `field.getGroupSettings()` (a
`PivotFieldGroupSettings`).

## Custom discrete groups (label items into buckets)

```java
PivotTable pt = ws.getPivotTables().get(0);
PivotField f = pt.getRowFields().get(0);
PivotFieldGroupOptions grp = new PivotFieldGroupOptions();
grp.setDiscreteGroupItems(new DiscreteGroupItem[]{
        new DiscreteGroupItem("H1", new int[]{1, 2, 3, 4, 5, 6}),   // months 1-6
        new DiscreteGroupItem("H2", new int[]{7, 8, 9, 10, 11, 12})});
f.group(grp);
pt.calculateData();
```

## Pitfalls

- **`groupBy(...)` and the table-level helpers are gone or dying** —
  `PivotField.groupBy(...)` overloads are deprecated (removal began
  Sept 2026) and the `PivotTable`-level auto/manual group-and-ungroup
  helpers were removed in 26.9. Use `group(...)` / `ungroup()` on the
  field.
- **Grouping is cache-level state** — always `calculateData()` /
  `refresh()` after grouping or ungrouping, before reading pivot cells
  or saving.
- **Group a field that is in an area** — the field must already be a
  row/column/page field; grouping a bare source column does nothing.

## Related

- [create-pivot-table.md](create-pivot-table.md)
- [pivot-fields-formatting.md](pivot-fields-formatting.md)
- [refresh-pivot-table.md](refresh-pivot-table.md)
