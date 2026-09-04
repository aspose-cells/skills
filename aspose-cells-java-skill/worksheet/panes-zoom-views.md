---
name: panes-zoom-views
description: Freeze / split panes, set zoom factor, show / hide gridlines, headers, formula bar, and tab strips in Aspose.Cells for Java. Use when the user asks to freeze header rows, lock a vertical pane, hide the gridlines, or set a custom zoom level.
applies_to: Java
keywords: [freeze pane, split, zoom, gridlines, view]
---

# Panes, zoom & view flags

## Imports

```java
import com.aspose.cells.*;
```

## Freeze pane

```java
Worksheet ws = wb.getWorksheets().get(0);
ws.freezePanes(2, 1, 2, 1);                     // row index, column index, total rows, total cols
```

Parameters: row/col index **to freeze at** in zero-based; the trailing
two arguments are usually the same as the leading two.

For Excel's "freeze top row": `freezePanes(1, 0, 1, 0)`.
For "freeze first column": `freezePanes(0, 1, 0, 1)`.

## Split panes (drag-style split)

```java
ws.setActiveCell("B5");
ws.split();
```

The split is positioned at the active cell when called.

## Zoom & view

```java
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().get(0);
Cells cells = ws.getCells();                   // unused; satisfies the wrapper
ws.setZoom(90);                                 // 90 %
ws.setTabColor(com.aspose.cells.Color.getSkyBlue());
ws.setSelected(true);                           // mark as selected on open
```

## Show / hide UI elements (workbook-scoped)

```java
WorkbookSettings settings = wb.getSettings();
settings.setShowTabs(false);                   // hide tabs strip
settings.setHScrollBarVisible(true);
settings.setVScrollBarVisible(true);

// Per-sheet gridlines / row-column headers are on Worksheet:
Worksheet ws = wb.getWorksheets().get(0);
ws.setGridlinesVisible(false);
ws.setRowColumnHeadersVisible(false);
```

## Page break preview

```java
ws.setType(SheetType.WORKSHEET);                // default
```

`Worksheet.setViewType(ViewType.PAGE_BREAK_PREVIEW)` puts the sheet
into the print-preview look — useful when generating PDFs.

## Pitfalls

- **`freezePanes` arguments** differ by overload: prefer the 4-int
  version with row/col + total rows/total cols, both trailing args
  matching the leading.
- **`setShowTabs(false)` doesn't hide the file** — Excel still opens
  it. Users just lose the UI affordance.
- **Zoom factor** is `int` percentage (10–400).
- **Scroll bars** settings persist to the workbook, not the sheet.

## Related

- [manage-worksheets.md](manage-worksheets.md)
- [page-setup-print.md](page-setup-print.md)
