# Styles and Formatting

Read this when formatting cells (font, color, borders, alignment, fill, number/date formats, conditional formatting, merged cells), a style change does not stick, dates display as serial numbers, or a save fails with "too many cell formats".

Core model: a cell's **value** and its **display format** are independent, and a cell's `Style` is a detached copy. Every visual change is the same shape: get or create a `Style`, mutate it, write it back. Colors are `aspose.pydrawing.Color` (`import aspose.pydrawing as drawing`).

## Pitfalls

### Style edits silently lost
Symptom: you write `cell.get_style().font.is_bold = True` (or change a color, border, or number format) and the saved file shows no change.
Cause: `Cell.get_style()` returns a **copy** of the cell's style, not a live reference. Mutating that copy changes nothing until you assign it back. There is no live `Cell.style` property.
Fix: get, mutate, then `set_style`.

```python
import aspose.cells as gc
import aspose.pydrawing as drawing

wb = gc.Workbook()
cell = wb.worksheets[0].cells.get("A1")
cell.get_style().font.is_bold = True   # WRONG: mutates a discarded copy
s = cell.get_style()
s.font.is_bold = True
s.font.color = drawing.Color.red
cell.set_style(s)                      # RIGHT: persists the change
```

### Per-cell styling in loops is slow and misscaled
Symptom: formatting thousands of cells one `get_style`/`set_style` at a time is slow and bloats the file.
Cause: styling cell-by-cell allocates and pool-compares a style per cell. When a whole row, column, or range shares one look, apply it once.
Fix: build one `Style` with `Workbook.create_style()`, declare a `StyleFlag` selecting which attributes to push, then apply at row/column/range scope. The `StyleFlag` is essential: only flagged attributes are written, so unrelated existing formatting on those cells is preserved.

```python
import aspose.cells as gc
import aspose.pydrawing as drawing

wb = gc.Workbook()
sheet = wb.worksheets[0]
st = wb.create_style()
st.font.is_bold = True
st.foreground_color = drawing.Color.light_gray
st.pattern = gc.BackgroundType.SOLID
st.number = 14
flag = gc.StyleFlag()
flag.font_bold = True
flag.cell_shading = True
flag.number_format = True
sheet.cells.apply_column_style(0, st, flag)   # whole column A
sheet.cells.apply_row_style(0, st, flag)      # whole row 1
```

Apply methods, each taking `(Style, StyleFlag)` plus an index where noted:
- `Cells.apply_column_style(column, style, flag)` - one column
- `Cells.apply_row_style(row, style, flag)` - one row
- `Cells.apply_style(style, flag)` - the entire worksheet
- `Range.apply_style(style, flag)` - a range (below)
- `Row.apply_style(style, flag)`, `Column.apply_style(style, flag)` - a row/column object

Verified `StyleFlag` bools include `all`, `font`, `font_bold`, `number_format`, `borders`, `cell_shading` (fill), `horizontal_alignment`, `wrap_text`. Each style attribute has a matching flag; set `all = True` to push everything.

```python
import aspose.cells as gc

wb = gc.Workbook()
sheet = wb.worksheets[0]
st = wb.create_style()
st.horizontal_alignment = gc.TextAlignmentType.CENTER
rng = sheet.cells.create_range(0, 0, 1, 4)   # A1:D1
flag = gc.StyleFlag()
flag.horizontal_alignment = True              # set StyleFlag flags as attributes, not kwargs
rng.apply_style(st, flag)
```

### "Too many cell formats" or a corrupted file
Symptom: save throws, or Excel reports the file is corrupt / hits a format limit, after formatting many cells.
Cause: allocating a fresh `Style` per cell inflates the workbook's unique-format table. Xlsx caps unique cell formats near **64,000**; the older `.xls` binary format allows far fewer (about 4,000). Exceed it and the file is invalid.
Fix: create a small set of reusable `Style` objects and share them across cells via `set_style` or the bulk apply methods above. `set_style` pools identical styles internally, so reusing one configured object is cheap. To build a style with no `Workbook` in hand, use `CellsFactory`:

```python
import aspose.cells as gc

factory = gc.CellsFactory()
s = factory.create_style()
s.font.is_bold = True
```

### Dates and numbers display as raw serial numbers
Symptom: a cell shows `45658` instead of `2025-01-01`, or `0.25` instead of `25%`.
Cause: value and format are independent. Storing a `datetime` or a fraction sets only the value; the cell keeps its `General` format and renders the underlying serial number.
Fix: put the value, then set a format on the style - a built-in id via `Style.number`, or a pattern via `Style.custom`.

```python
import aspose.cells as gc
from datetime import datetime

wb = gc.Workbook()
cell = wb.worksheets[0].cells.get("A1")
cell.put_value(datetime(2025, 1, 1))   # value only; still shows a serial number
s = cell.get_style()
s.number = 14               # built-in short date; OR:
s.custom = "yyyy-mm-dd"     # custom pattern (overrides number)
cell.set_style(s)
```

`number` and `custom` override each other - setting one clears the other. Read a formatted date back as a typed value with `Cell.date_time_value` (see cell-values-and-data.md).

Useful built-in `Style.number` ids:

| id | Type | Format |
|---:|------|--------|
| 0 | General | General |
| 1 | Decimal | 0 |
| 2 | Decimal | 0.00 |
| 4 | Decimal | #,##0.00 |
| 7 | Currency | $#,##0.00 |
| 9 | Percentage | 0% |
| 10 | Percentage | 0.00% |
| 14 | Date | short date (m/d/yyyy) |
| 15 | Date | d-mmm-yy |
| 21 | Time | h:mm:ss |
| 22 | Date-time | m/d/yy h:mm |
| 49 | Text | @ |

For anything outside this set (e.g. `yyyy-mm-dd hh:mm`, or a currency suffix), use `Style.custom`.

### 1904 date system shifts dates by ~4 years
Symptom: dates are off by roughly four years (1462 days) when interoperating with files authored in classic Mac Excel.
Cause: workbooks use either the 1900 date system (Windows default) or the 1904 system (old Mac). The same serial number resolves to a different date under each.
Fix: match the workbook to the file's system with `Workbook.settings.date1904`.

```python
import aspose.cells as gc

wb = gc.Workbook()
wb.settings.date1904 = True   # interpret and store serials in the 1904 system
```

## Fonts, colors, borders, fill, alignment

All live on `Style`; set them, then `set_style` (or apply in bulk). Key members:
- Font: `style.font.name`, `.size`, `.is_bold`, `.is_italic`, `.underline`, `.color`.
- Fill: `style.foreground_color` is the fill color, but it only shows once `style.pattern` is set. Use `style.pattern = gc.BackgroundType.SOLID` for a flat fill; other `BackgroundType` values give patterns, where `style.background_color` is the second color.
- Alignment: `style.horizontal_alignment` / `style.vertical_alignment = gc.TextAlignmentType.CENTER` (also `LEFT`, `RIGHT`, `TOP`, `BOTTOM`, `JUSTIFY`); `style.is_text_wrapped = True` to wrap.
- Borders: `style.set_border(BorderType, CellBorderType, Color)` per edge, or `style.borders[BorderType.TOP_BORDER].line_style = CellBorderType.THIN` with `.color`. Always set line style and color together.

```python
import aspose.cells as gc
import aspose.pydrawing as drawing

wb = gc.Workbook()
cell = wb.worksheets[0].cells.get("B2")
s = cell.get_style()
s.font.name = "Calibri"
s.font.size = 12
s.font.color = drawing.Color.white
s.foreground_color = drawing.Color.steel_blue   # fill color...
s.pattern = gc.BackgroundType.SOLID            # ...needs a pattern to show
s.horizontal_alignment = gc.TextAlignmentType.CENTER
s.is_text_wrapped = True
s.set_border(gc.BorderType.BOTTOM_BORDER, gc.CellBorderType.THIN, drawing.Color.black)
cell.set_style(s)
```

Foreground vs background gotcha: if `pattern` is `NONE`, `foreground_color` is ignored and reads back as `Color.empty`; `background_color` applies only for true patterns, not for `SOLID`. For a plain fill, set `foreground_color` + `SOLID`.

## Built-in styles

`Workbook.create_builtin_style(BuiltinStyleType)` returns a ready-made named style (`TITLE`, `GOOD`, `BAD`, `NEUTRAL`, `NOTE`, `TOTAL`, plus accent and percent variants). `Workbook.get_named_style(name)` fetches a named style already in the pool.

```python
import aspose.cells as gc

wb = gc.Workbook()
title_style = wb.create_builtin_style(gc.BuiltinStyleType.TITLE)
cell = wb.worksheets[0].cells.get("A1")
cell.put_value("Report")
cell.set_style(title_style)
```

## Conditional formatting

Rules live per worksheet. Add a formatting block, attach one or more cell areas, add conditions, then style each condition. Only font (color/bold/underline/strikeout), borders, and pattern can change conditionally.

```python
import aspose.cells as gc
import aspose.pydrawing as drawing

wb = gc.Workbook()
sheet = wb.worksheets[0]
idx = sheet.conditional_formattings.add()
fcc = sheet.conditional_formattings[idx]
fcc.add_area(gc.CellArea.create_cell_area("A1", "A10"))
ci = fcc.add_condition(gc.FormatConditionType.CELL_VALUE, gc.OperatorType.GREATER_THAN, "100", None)
# CF fills use background_color - a deliberate exception to the Foreground+Solid rule above
fcc[ci].style.background_color = drawing.Color.red   # applied where A1:A10 > 100
```

- `FormatConditionType`: `CELL_VALUE`, `EXPRESSION` (formula-driven), `COLOR_SCALE`, `DATA_BAR`, `ICON_SET`, and more.
- `OperatorType`: `BETWEEN`, `NOT_BETWEEN`, `GREATER_THAN`, `LESS_THAN`, `GREATER_OR_EQUAL`, `EQUAL`, and so on. For `BETWEEN` pass both `formula1` and `formula2`; for one-sided operators pass `None` as the last argument.
- For a formula rule use `FormatConditionType.EXPRESSION` and put the formula in `formula1` (e.g. `"=MOD(ROW(),2)=0"` for alternate-row shading).

Styling pivot table cells is owned by pivot-tables.md.

## Merged cells

`Cells.merge(first_row, first_column, total_rows, total_columns)` merges; `Cells.un_merge(...)` with the same coordinates splits. Only the top-left cell's value and format survive a merge - data in the other cells is discarded.

```python
import aspose.cells as gc

wb = gc.Workbook()
cells = wb.worksheets[0].cells
cells.merge(0, 0, 2, 3)            # merge A1:C2 (2 rows, 3 columns)
cells.get("A1").put_value("Merged title")
```

Merged regions interact with autofit and row-height calculations; that interplay is owned by worksheets-rows-columns.md.

## Related
- cell-values-and-data.md - reading typed values (date_time_value), put_value, list import/export
- worksheets-rows-columns.md - autofit, row height, ranges, merged-cell layout
- pivot-tables.md - styling pivot table cells
- formulas-and-calculation.md - recalculating before reading computed values
