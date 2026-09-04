---
name: manage-worksheets
description: Add, rename, remove, copy, reorder, and set tab colour for worksheets in Aspose.Cells for Java. Use when the user asks to create new sheets, swap sheet order, clone a sheet, or hide a tab.
applies_to: Java
keywords: [worksheet, sheet, add, rename, copy, remove, reorder, tab color]
---

# Manage worksheets

## Imports

```java
import com.aspose.cells.*;
```

## Add / remove / rename

```java
WorksheetCollection sheets = wb.getWorksheets();

// Add a new sheet, choose position
Worksheet report = sheets.add("Report");
Worksheet backup = sheets.add("Backup");     // worksheet type is the default

// Rename
report.setName("Daily Report");

// Remove by index
sheets.removeAt(0);

// Active sheet (Excel opens with this one visible)
wb.getWorksheets().setActiveSheetIndex(1);
```

## Copy a sheet

```java
WorksheetCollection sheets = wb.getWorksheets();
Worksheet report = sheets.get("Report");
Worksheet copy = sheets.add("Copy of Q1");
copy.copy(report);
```

## Move (re-order)

```java
WorksheetCollection sheets = wb.getWorksheets();
sheets.get("Backup").moveTo(0);               // become first
```

`moveTo` re-positions the sheet within the collection, including
across `Workbook`s? — no; only within the same workbook.

## Hide / show & tab colour

```java
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().get(0);
WorksheetCollection sheets = wb.getWorksheets();
Worksheet report = sheets.add("Report");
Worksheet backup = sheets.add("Backup");
Cells cells = ws.getCells();
backup.setVisible(false);                       // fully hidden
// for VERY_HIDDEN use backup.setVisibilityType(VisibilityType.VERY_HIDDEN);

report.setTabColor(com.aspose.cells.Color.getCoral());
```

## Select active sheet at runtime

```java
WorksheetCollection sheets = wb.getWorksheets();
wb.getWorksheets().setActiveSheetIndex(sheets.get("Report").getIndex());
```

## Pitfalls

- **`setName` rejects duplicate names** — even case-insensitive
  ("Report" vs "REPORT") within the same workbook. Renaming produces
  a runtime exception.
- **Removing the active sheet** promotes another sheet automatically
  — re-check `setActiveSheetIndex`.
- **Cross-workbook copies** are not supported on a single call —
  build the destination workbook fresh and copy sheets in.
- **`VERY_HIDDEN` (via `setVisibilityType(VisibilityType.VERY_HIDDEN)`) is
  intentionally obscure** — it is only hidden from the UI while the workbook
  is unprotected; call `ws.setVisible(false)` plus `wb.protect(...)` so users
  cannot re-show the sheet from the workbook manager.

## Related

- [page-setup-print.md](page-setup-print.md)
- [panes-zoom-views.md](panes-zoom-views.md)
