---
name: cell-formatting-styles
description: Apply fonts, colours, borders, fills, alignment, number formats, and named styles to cells in Aspose.Cells for Java. Use whenever the user asks to format, colour, bold, or otherwise style cell content.
applies_to: Java
keywords: [style, format, font, color, border, fill, alignment, number, style]
---

# Cell formatting & styles

## Imports

```java
import com.aspose.cells.*;
```

## Quick example — single cell

```java
Cell c = ws.getCells().get("A1");

Style s = c.getStyle();
s.getFont().setBold(true);
s.getFont().setColor(com.aspose.cells.Color.getRed());
s.setPattern(BackgroundType.SOLID);
s.setForegroundColor(com.aspose.cells.Color.getLightGray());
c.setStyle(s);
```

## Bulk — apply to a range

```java
Style s = wb.createStyle();
s.getFont().setName("Calibri");
s.getFont().setSize(11);
s.setHorizontalAlignment(TextAlignmentType.CENTER);

StyleFlag flag = new StyleFlag();
flag.setFont(true);
flag.setAlignments(true);

ws.getCells().createRange("A1:Z1").applyStyle(s, flag);
```

Use `StyleFlag` to choose which sub-fields are applied — copying every
field can clobber existing per-cell data formats.

## Number / date formats

```java
Cell cell = ws.getCells().get("B2");
Style s = cell.getStyle();
s.setNumber(2);                                  // built-in
s.setCustom("0.00%");                            // custom
s.setCustom("yyyy-MM-dd");                       // date
cell.setStyle(s);
```

`Style.setCustom(...)` overrides `setNumber`.

## Borders

```java
Cell cell = ws.getCells().get("C3");
Style s = cell.getStyle();
BorderCollection borders = s.getBorders();
borders.getByBorderType(BorderType.TOP_BORDER)
       .setLineStyle(CellBorderType.THIN);
borders.getByBorderType(BorderType.TOP_BORDER)
       .setColor(com.aspose.cells.Color.getBlack());
cell.setStyle(s);
```

## Reuse style efficiently

Modifying `cell.getStyle()` then `setStyle(s)` creates a duplicate
style object — many duplicates slow large workbooks down. To reuse a
shared `Style`:

```java
Style shared = wb.createStyle();                      // new shared Style
shared.getFont().setBold(true);

Cell c = ws.getCells().get("A1");
c.setStyle(shared);
```

## Pitfalls

- **Don't forget `StyleFlag`** when using `applyStyle` — without it,
  every flag is true and you'll overwrite number formats and borders.
- **`getStyle()` is a copy** — never call it inside a row loop on a
  million-row sheet, that allocates a copy per cell.
- **Color values** are `java.awt.Color`; for theme colours use
  `wb.getThemeColor(ThemeColorType.ACCENT_1)` or pass
  `Color.fromArgb(0xFF, 0x33, 0x00)`.
- **Excel `CellBorderType.DOUBLE` writes a different attribute** than
  `THICK` — if the user said "double border", use DOUBLE, not THICK.
- **Pattern / gradient** are mutually exclusive on a single cell.

### "Too many cell formats" or a corrupted file

Symptom: save throws, or Excel reports the file is corrupt / hits a
format limit, after formatting many cells.

Cause: allocating a fresh `Style` per cell inflates the workbook's
unique-format table. XLSX caps unique cell formats near **64,000**;
the older `.xls` binary format allows far fewer (~4,000). Exceed
the cap and the file is invalid.

Fix: create a small set of reusable `Style` objects and share them
across cells via `setStyle`, or apply them in bulk via the
`applyColumnStyle` / `applyRowStyle` / `applyStyle` methods. These
apply the same `Style` reference to many cells so the unique-format
table stays small.

```java
Style s = wb.createStyle();
s.getFont().setBold(true);
s.setPattern(BackgroundType.SOLID);
s.setForegroundColor(com.aspose.cells.Color.getLightGray());

StyleFlag flag = new StyleFlag();
flag.setFont(true);
flag.setCellShading(true);

cells.applyColumnStyle(0, s, flag);   // whole column A
cells.applyRowStyle(0, s, flag);      // whole row 1
cells.applyStyle(s, flag);            // entire worksheet
cells.createRange("A1:D10").applyStyle(s, flag);  // a rectangular range
```

To build a `Style` with no `Workbook` in hand, use `CellsFactory`:

```java
com.aspose.cells.CellsFactory factory = new com.aspose.cells.CellsFactory();
Style s = factory.createStyle();
s.getFont().setBold(true);
```

### `getStyle()` returns a copy — mutating it does not stick

Symptom: `cell.getStyle().getFont().setBold(true)` (or any other
mutation) appears to do nothing — the saved file has no change.

Cause: `Cell.getStyle()` returns a **copy** of the cell's style,
not a live reference. Mutating the copy has no effect until you
assign it back. There is no live `cell.getStyle()` setter; the
property was removed in 7.1.0.

Fix: get, mutate, then `setStyle`.

```java
Cell cell = ws.getCells().get("A1");
Style s = cell.getStyle();                  // copy
s.getFont().setBold(true);
s.getFont().setColor(com.aspose.cells.Color.getRed());
cell.setStyle(s);                           // assign back
```

## Related

- [conditional-formatting.md](conditional-formatting.md)
- [rows-columns-operations.md](rows-columns-operations.md)
