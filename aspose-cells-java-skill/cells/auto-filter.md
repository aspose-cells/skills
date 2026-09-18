---
name: auto-filter
description: Add Excel-style AutoFilter dropdowns to ranges in Aspose.Cells for Java — including custom criteria, top-N filters, blank handling, and refreshing visible rows after filter changes.
applies_to: Java
keywords: [autofilter, filter, dropdown, criteria]
---

# Auto filter

## Imports

```java
import com.aspose.cells.*;
```

## Quick example — enable on a range

```java
Worksheet ws = wb.getWorksheets().get(0);
ws.getAutoFilter().setRange("A1:E100");
```

## Apply filter criteria

```java
AutoFilter af = ws.getAutoFilter();
af.filterTop10(0, true, true, 5);                  // top 5 of column A by value
af.addFilter(0, "Closed");                         // exact value match on column A
af.addFilter(0, "Pending");
af.custom(0, FilterOperatorType.EQUAL, "Closed");  // operator-based match
```

After mutating criteria, call `af.refresh()` to update which rows are
visible.

## Read filtered rows

```java
int total = ws.getCells().getMaxDataRow();
for (int r = 1; r <= total; r++) {
    Row row = ws.getCells().getRows().get(r);
    if (!row.isHidden()) {
        // process the row
    }
}
```

## Pitfalls

- **`setRange` re-anchors the autofilter**. Do it **before** adding
  criteria.
- **`refresh()` is mandatory** — without it, `isHidden()` is
  stale.
- **Combined criteria** (AND / OR) — use the `addFilter(...)` overload
  with a list, or set `custom(...)` with the desired operator.
- **Top-10 ranking** is inclusive: `top10` of 19 rows returns 2; for
  "largest 10 %" pass `topN = 10` and `isPercent = true`.

## Related

- [sort-data.md](sort-data.md)
- [data-validation.md](data-validation.md)
