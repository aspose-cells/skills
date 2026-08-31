# Worksheets, Rows, Columns, Ranges, Tables, Protection

Read this when adding/copying/moving/deleting sheets, inserting/deleting/hiding rows or columns, autofitting, named ranges, Excel tables (ListObjects), data validation, freeze panes, or protecting workbooks/worksheets/cells.

All indices are 0-based: row 0 is Excel row 1, column 0 is column A. The workhorse is `Worksheet.cells`.

## Pitfalls

### Off-by-one and self-shifting indices when inserting/deleting

Symptom: content lands one row/column off, or a delete-in-a-loop removes the wrong rows and skips some.

Cause: indices are 0-based (Excel row 1 = index 0). `insert_rows`/`delete_rows` shift ALL existing content and auto-update formulas, chart sources, and references; deleting index `r` renumbers every row below it, so a forward loop targets the wrong rows after the first pass.

Fix: never manually re-adjust references after an insert/delete - Aspose does it for you. Delete a contiguous block in ONE call (`cells.delete_rows(2, 5)`). For scattered rows, iterate descending so earlier deletes cannot shift indices you have not processed yet.

```python
import aspose.cells as gc

wb = gc.Workbook("input.xlsx")
cells = wb.worksheets[0].cells
# Delete scattered rows 2, 5, 9 (0-based) top-safe: process DESCENDING.
for r in [9, 5, 2]:
    cells.delete_rows(r, 1)
wb.save("output.xlsx")
```

`insert_rows(row_index, total_rows)` inserts above `row_index` and pushes content down; `insert_column(column_index)` pushes right.

### AutoFit is slow or "does nothing"

Symptom: `auto_fit_columns`/`auto_fit_rows` takes seconds on large sheets, or leaves merged/wrapped cells unsized.

Cause: autofit measures every cell in scope - it is O(cells) and expensive. By design it ignores merged cells and wrapped-text columns unless you pass options.

Fix: call it ONCE after all data is written, never per-row. Narrow the scope with `auto_fit_columns(first_column, last_column)`. For merged cells pass `AutoFitterOptions`.

```python
import aspose.cells as gc

wb = gc.Workbook("input.xlsx")
sheet = wb.worksheets[0]
# ... write ALL data first, then autofit once ...
opts = gc.AutoFitterOptions()
opts.auto_fit_merged_cells_type = gc.AutoFitMergedCellsType.EACH_LINE
opts.only_auto = True              # skip rows/cols the user sized manually
sheet.auto_fit_columns(opts)
sheet.auto_fit_rows(opts)
wb.save("output.xlsx")
```

`AutoFitMergedCellsType`: `NONE`, `FIRST_LINE`, `LAST_LINE`, `EACH_LINE`. Other options: `ignore_hidden`, `max_row_height`, `auto_fit_wrapped_text_type`.

### FreezePanes argument order

Symptom: the wrong rows freeze, or nothing looks frozen.

Cause: misreading `freeze_panes(int row, int column, int freezed_rows, int freezed_columns)`. The first pair is the top-left cell of the SCROLLABLE area (where the split sits); the second pair is HOW MANY rows/columns stay frozen.

Fix: to freeze only the top row, split at A2 (row 1, col 0) and freeze 1 row, 0 columns.

```python
import aspose.cells as gc

wb = gc.Workbook("input.xlsx")
sheet = wb.worksheets[0]
sheet.freeze_panes(1, 0, 1, 0)     # freeze the top row
# Freeze first two rows AND first column: freeze_panes(2, 1, 2, 1)
wb.save("output.xlsx")
```

`freeze_panes(string cell_name, int freezed_rows, int freezed_columns)` is an equivalent overload.

## Worksheet lifecycle

`Workbook.worksheets` is the `WorksheetCollection`. `add(name)` returns the `Worksheet`; `add()` returns an int index. Access by name or index; remove with `remove_at`.

```python
import aspose.cells as gc

wb = gc.Workbook()
data = wb.worksheets.add("Data")   # returns the Worksheet
by_name = wb.worksheets.get("Sheet1")
by_index = wb.worksheets[0]
wb.worksheets.active_sheet_index = data.index   # sheet selected on open
wb.worksheets.remove_at("Sheet1")               # remove_at takes the sheet name (str) in this binding
wb.save("output.xlsx")
```

Duplicate within a workbook with `add_copy` (by index or name); reorder with `move_to`; copy across workbooks with `dest.copy(source)`.

```python
import aspose.cells as gc

wb = gc.Workbook("input.xlsx")
copy_index = wb.worksheets.add_copy(0)      # duplicate sheet 0
wb.worksheets[copy_index].move_to(0)        # move the copy to the front
other = gc.Workbook()
other.worksheets[0].copy(wb.worksheets[0])   # cross-workbook: dest.copy(source)
other.save("output.xlsx")
```

Tab color and visibility:

```python
import aspose.cells as gc
import aspose.pydrawing as drawing

wb = gc.Workbook("input.xlsx")
sheet = wb.worksheets[0]
sheet.tab_color = drawing.Color.green      # aspose.pydrawing.Color
sheet.is_visible = False                   # hide sheet (at least one must stay visible)
wb.settings.show_tabs = True                # toggle the whole tab bar
wb.save("output.xlsx")
```

## Rows and columns

Insert, hide, group, and size rows/columns through `cells`. Column width is in characters; row height is in points.

```python
import aspose.cells as gc

wb = gc.Workbook("input.xlsx")
cells = wb.worksheets[0].cells
cells.insert_rows(2, 3)            # 3 rows above index 2
cells.insert_column(1)            # 1 column at index 1
cells.hide_row(3)
cells.hide_columns(1, 2)          # hide 2 columns starting at index 1
cells.group_rows(5, 9, False)     # outline rows 5-9, not collapsed
cells.set_column_width(0, 25.0)   # characters (float, not int)
cells.set_row_height(0, 30.0)     # points (float, not int)
wb.save("output.xlsx")
```

Reveal hidden rows/columns with `unhide_row(index, height)` / `unhide_column(index, width)` (pass a negative width to restore the prior width). Ungroup with `ungroup_rows`/`ungroup_columns`. Outline direction is `Worksheet.outline.summary_row_below` / `summary_column_right`.

## Ranges

Create a range by address or by `(first_row, first_column, total_rows, total_columns)`. The range indexer is 0-based WITHIN the range. Set `Range.name` to make it a named range; fetch with `Worksheets.get_range_by_name`.

```python
import aspose.cells as gc

wb = gc.Workbook()
cells = wb.worksheets[0].cells
rng = cells.create_range(0, 0, 4, 4)  # A1:D4
# Range is NOT subscriptable; address cells via the worksheet using the range origin
cells.get(rng.first_row, rng.first_column).put_value("top-left")
rng.name = "MyData"               # promote to a named range
named = wb.worksheets.get_range_by_name("MyData")
wb.save("output.xlsx")
```

Copy a range with `copy` (all content) or `copy(source, PasteOptions)` for paste-special:

```python
import aspose.cells as gc

wb = gc.Workbook("input.xlsx")
cells = wb.worksheets[0].cells
source = cells.create_range("A1:C5")
target = cells.create_range("E1:G5")
target.copy(source)                # values, formulas, formats, drawings
opts = gc.PasteOptions()
opts.paste_type = gc.PasteType.VALUES   # or FORMATS, FORMULAS, VALUES_AND_FORMATS...
target.copy(source, opts)
wb.save("output.xlsx")
```

Workbook-vs-worksheet-scoped defined names and the `names` collection -> formulas-and-calculation.md. Merge/unmerge and range styling -> styles-and-formatting.md.

## Tables (ListObjects)

`Worksheet.list_objects.add(start_row, start_col, end_row, end_col, has_headers)` returns the new table's index; index into `list_objects` to configure it.

```python
import aspose.cells as gc

wb = gc.Workbook("input.xlsx")
sheet = wb.worksheets[0]
idx = sheet.list_objects.add(0, 0, 9, 3, True)   # A1:D10, first row = header
table = sheet.list_objects[idx]
table.table_style_type = gc.tables.TableStyleType.TABLE_STYLE_MEDIUM9
table.show_totals = True
wb.save("output.xlsx")
```

Convert a table back to plain cells with `table.convert_to_range()` (formatting is kept, but table behavior and structured references are dropped).

## Data validation

Each `Worksheet.validations` entry covers one or more `CellArea`s. `add(CellArea)` returns the index; configure the returned `Validation`. For a dropdown, `in_cell_drop_down` must be true.

```python
import aspose.cells as gc

wb = gc.Workbook()
sheet = wb.worksheets[0]
area = gc.CellArea.create_cell_area("A1", "A10")
v = sheet.validations[sheet.validations.add(area)]
v.type = gc.ValidationType.LIST
v.in_cell_drop_down = True
v.formula1 = "Red,Green,Blue"      # inline list, or "=Sheet2!$A$1:$A$3"
wb.save("output.xlsx")
```

Numeric range with input/error prompts:

```python
import aspose.cells as gc

wb = gc.Workbook()
sheet = wb.worksheets[0]
area = gc.CellArea.create_cell_area("B1", "B1")
v = sheet.validations[sheet.validations.add(area)]
v.type = gc.ValidationType.WHOLE_NUMBER    # or DECIMAL, DATE, TIME, TEXT_LENGTH
v.operator = gc.OperatorType.BETWEEN
v.formula1 = "1"
v.formula2 = "100"
v.alert_style = gc.ValidationAlertType.STOP   # STOP, WARNING, INFORMATION
v.error_message = "Enter a whole number from 1 to 100."
v.show_error = True
wb.save("output.xlsx")
```

Add another region to an existing rule with `Validation.add_area(cell_area)`. Types: `ANY_VALUE`, `WHOLE_NUMBER`, `DECIMAL`, `LIST`, `DATE`, `TIME`, `TEXT_LENGTH`, `CUSTOM`.

## Protection

Worksheet protection restricts editing but does NOT encrypt data. Cells are locked by default; the classic pattern is to unlock only the cells that should stay editable, then protect. `StyleFlag.locked = True` tells `apply_style` to apply just the lock bit.

```python
import aspose.cells as gc

wb = gc.Workbook("input.xlsx")
sheet = wb.worksheets[0]
style = wb.create_style()
style.is_locked = False
flag = gc.StyleFlag()
flag.locked = True
editable = sheet.cells.create_range("B2", "C3")
editable.apply_style(style, flag)       # unlock B2:C3
sheet.protect(gc.ProtectionType.ALL)    # lock everything else
wb.save("output.xlsx")
```

Use `protect(ProtectionType)` or `protect(ProtectionType, new_password, old_password)`. `ProtectionType`: `ALL`, `CONTENTS`, `OBJECTS`, `SCENARIOS`, `STRUCTURE`, `WINDOWS`, `NONE`. To lock a whole row/column, use `Row.apply_style` / `Column.apply_style` instead of a range.

Workbook structure protection (blocks add/move/delete/rename of sheets) and file open-password encryption are separate operations:

```python
import aspose.cells as gc

wb = gc.Workbook("input.xlsx")
wb.protect(gc.ProtectionType.STRUCTURE, "structPwd")   # structure lock
wb.settings.password = "openPwd"                       # open-password encryption
wb.save("output.xlsx")                                # XLSX -> AES; XLS -> weaker
```

`Workbook.settings.password` sets the file open password (whole-file encryption); it applies to XLSX/XLS on save. Clear structure protection with `Workbook.unprotect(password)`.

## Grid limits

XLS grids are 65536 x 256; XLSX/XLSM are 1048576 x 16384. Writing beyond the target format's limits fails or truncates - pick the format accordingly (format trade-offs -> conversion-and-rendering.md).

## Related
- styles-and-formatting.md - cell/row/column style application, merged-cell formatting, borders.
- cell-values-and-data.md - reading/writing values and iterating cells efficiently.
- formulas-and-calculation.md - defined names, named-range scope, formula recalculation.
- performance-and-large-files.md - large-sheet memory and speed, LightCells.
- conversion-and-rendering.md - format choice, grid-size limits, save options.
