# Cell Values and Data IO

Read this when setting or reading cell values, getting unexpected runtime types or
formatted strings on read, importing/exporting lists/objects, iterating
cells, doing find/replace, or adding hyperlinks.

All row/column indices are 0-based (row 0 = Excel row 1). Import and `put_value` write
VALUES ONLY - number/date display formats are a separate step; see
styles-and-formatting.md. Get a cell with `cells.get("A1")` or `cells.get(row, col)`.
There is no `cells["A1"]` indexer in Python.

## Pitfalls

### Cell.value returns an unpredictable boxed type

Symptom: `cell.value` (or a typed accessor) surprises you; a date cell yields a
number; an integer sometimes boxes as `int`, sometimes as `float`.
Cause: `Cell.value` is `object`. Its runtime type is one of `None`, `bool`, `datetime`,
`float`, `int`, `str` - and Aspose does not guarantee a stable type (an integer may
box as `int` or `float` across calls/versions). Dates are stored as OADate serials, so
a date cell's `value` may box as `datetime` OR surface as a `float`.
Fix: never cast `Cell.value` directly. Branch on `Cell.type` (a `CellValueType`) and
read through the typed accessors, which coerce reliably.

| Accessor | Returns | Use when |
|---|---|---|
| `Cell.value` | `object` (unstable type) | you will type-check it yourself |
| `Cell.double_value` / `float_value` / `int_value` | number | `type == IS_NUMERIC` |
| `Cell.date_time_value` | `datetime` | `type == IS_DATE_TIME` |
| `Cell.bool_value` | `bool` | `type == IS_BOOL` |
| `Cell.string_value` | formatted `str` | any type |

`CellValueType` members: `IS_NULL`, `IS_NUMERIC`, `IS_DATE_TIME`, `IS_STRING`, `IS_BOOL`,
`IS_ERROR`, `IS_UNKNOWN`. `double_value`/`float_value` raise on an empty (`IS_NULL`) cell -
guard first.

```python
import aspose.cells as gc

wb = gc.Workbook("input.xlsx")
cell = wb.worksheets[0].cells.get("A1")
if cell.type == gc.CellValueType.IS_NUMERIC:
    print(cell.double_value)
elif cell.type == gc.CellValueType.IS_DATE_TIME:
    print(cell.date_time_value)
elif cell.type == gc.CellValueType.IS_BOOL:
    print(cell.bool_value)
elif cell.type == gc.CellValueType.IS_STRING:
    print(cell.string_value)
elif cell.type == gc.CellValueType.IS_ERROR:
    print(cell.string_value)   # "#DIV/0!"
# IS_NULL -> empty, do nothing
```

### Read string differs from what Excel shows

Symptom: a cell displays `0.01` in Excel but your read returns `0.012345`, or vice versa.
Cause: three different strings exist for one cell. `string_value` applies the cell's own
number format (identical to Excel "copy as text"). `display_string_value` applies the
cell's display style - exactly what Excel renders (reflects conditional formatting).
The underlying unformatted value is neither.
Fix: pick the accessor that matches intent. For explicit control call
`get_string_value(CellValueFormatStrategy)`: `NONE` = unformatted, `CELL_STYLE` = cell's
format, `DISPLAY_STYLE` = as displayed.

```python
import aspose.cells as gc

wb = gc.Workbook("input.xlsx")
cell = wb.worksheets[0].cells.get("A1")              # 0.012345, number format "0.00"
print(cell.string_value)                            # "0.01"     - cell number format
print(cell.display_string_value)                    # "0.01"     - what Excel shows
print(cell.get_string_value(gc.CellValueFormatStrategy.NONE))  # "0.012345" - raw
```

### Text numbers stay text and break SUM

Symptom: imported/typed numbers are left-aligned, `SUM` ignores them, sorting is
lexical. Cause: a string like `"00123"` written as text stays text. `put_value(string)`
does NOT parse. Fix: parse on write with the `put_value(string, is_converted)` overload
(`True` converts to number/date/bool), or fix a whole sheet after load with
`Cells.convert_string_to_numeric_value()`.

```python
import aspose.cells as gc

wb = gc.Workbook()
cells = wb.worksheets[0].cells
cells.get("A1").put_value("00123")        # stays text -> left-aligned, excluded from SUM
cells.get("A2").put_value("00123", True)  # parsed to numeric 123
cells.convert_string_to_numeric_value()   # whole sheet: convert every text-number in place
```

### Scanning with cells.get(r, c) is slow and instantiates empty cells

Symptom: iterating a large or sparse sheet is slow or raises `OutOfMemoryError`.
Cause: `cells.get(r, c)` and `cells.get("A1")` CREATE a blank `Cell` when none exists, so
a full `r x c` scan instantiates the entire matrix. `max_data_row`/`max_data_column` also
recompute statistics on every access.
Fix: enumerate only the cells that exist with the `rows` -> `row` -> `cell` iterators.
Do not add or delete cells while enumerating (the enumerator may skip or repeat). If you
must loop by index, cache `max_data_row`/`max_data_column` once - both return **-1** on an
empty sheet, so `for r in range(max_row + 1)` correctly runs zero times.

```python
import aspose.cells as gc

wb = gc.Workbook("input.xlsx")
cells = wb.worksheets[0].cells
for row in cells.rows:                    # only rows that exist
    for cell in row:                      # only instantiated cells in that row
        if cell.type == gc.CellValueType.IS_NULL:
            continue
        print(f"{cell.row},{cell.column} = {cell.string_value}")
```

`Cells`, `Row`, and `Range` each expose iteration yielding `Cell`; `Row` also gives
`first_cell`, `last_cell`, `last_data_cell`, and `index`.

## Setting values

`Cell.put_value` has typed overloads - `put_value(float)`, `put_value(int)`,
`put_value(bool)`, `put_value(datetime)`, `put_value(str)`, plus `put_value(object)` and
the parsing `put_value(str, bool)`. Setting `Cell.value = x` is equivalent to the
matching typed overload. For bulk writes, populate first by row then by column - it is
markedly faster than column-major order.

```python
import aspose.cells as gc
from datetime import datetime

wb = gc.Workbook()
cells = wb.worksheets[0].cells
cells.get("A1").put_value("Total")     # string
cells.get("B1").put_value(42.5)        # float
cells.get("C1").put_value(True)        # bool
cells.get("D1").put_value(datetime.now())  # datetime - still needs a date format to display right
cells.get("E1").value = 7              # equivalent to put_value(7)
```

## Importing arrays and objects

`import_array` writes a 1D/2D `list` of **strings** `(array, first_row, first_column, is_vertical)`;
numeric values (`int`/`float`) are rejected (`TypeError: can't build String from 'float'`), so use
`import_object_array` for numbers. `import_object_array(object[, ])` lays a
rectangular block (the Python equivalent of .NET `ImportTwoDimensionArray`). For a list of
business objects, assemble a 2D grid and use `import_object_array` (the portable path);
`import_custom_objects(...)` exists but maps over .NET-reflectable types and is unreliable
for arbitrary Python objects in this binding.

```python
import aspose.cells as gc

wb = gc.Workbook()
cells = wb.worksheets[0].cells
header = ["Name", "Qty"]
cells.import_array(header, 0, 0, False)        # False = horizontal: A1, B1
prices = [1.5, 3.0, 2.25]
cells.import_object_array(prices, 1, 0, True)  # True = vertical: A2, A3, A4 (numbers need import_object_array)
grid = [["Pen", 12], ["Cup", 5]]
cells.import_object_array(grid, 1, 1, False)   # 2x2 block anchored at B2; False = horizontal
```

```python
import aspose.cells as gc

class Product:
    def __init__(self, name, price):
        self.name = name
        self.price = price

items = [Product("Pen", 1.5), Product("Cup", 3.0)]
wb = gc.Workbook()
cells = wb.worksheets[0].cells
# import_custom_objects over plain Python objects is unreliable in this binding
# (Aspose reflects over .NET types, not arbitrary Python attributes), so the
# portable path is a 2D grid via import_object_array:
grid = [["Name", "Price"]] + [[it.name, it.price] for it in items]
cells.import_object_array(grid, 0, 0, False)   # False = horizontal
```

## Exporting to a list

This binding does not expose the .NET `ExportArray` single call; read a rectangular block
by iterating cells instead. `Cell.value` returns the typed value (`None` for an empty
cell), and the typed views (`string_value`, `double_value`, `bool_value`, ...) give
per-type access (see "reading values" above).

```python
import aspose.cells as gc

wb = gc.Workbook("input.xlsx")
cells = wb.worksheets[0].cells
rows, cols = 10, 3
data = [
    [cells.get(r, c).value for c in range(cols)]
    for r in range(rows)
]
print(len(data), "rows")
```

## Find and replace

`Cells.find(object value, Cell previous_cell, FindOptions options)` returns the matching
`Cell` or `None`; pass the previous hit back in to continue. Set `FindOptions.look_in_type`
(`VALUES`, `FORMULAS`, `ORIGINAL_VALUES`, `FORMATTED_VALUES`, ...) and
`FindOptions.look_at_type` (`CONTAINS`, `START_WITH`, `END_WITH`, `ENTIRE_CONTENT`); other
options include `case_sensitive` and `regex_key`. For blanket substitution,
`Worksheet.replace(old_string, new_string)` returns the count of changed cells.

```python
import aspose.cells as gc

wb = gc.Workbook("input.xlsx")
cells = wb.worksheets[0].cells
opts = gc.FindOptions()
opts.look_in_type = gc.LookInType.VALUES
opts.look_at_type = gc.LookAtType.ENTIRE_CONTENT  # whole-cell match; CONTAINS for substring
hits = 0
found = cells.find("Oranges", None, opts)
while found is not None:
    hits += 1
    found = cells.find("Oranges", found, opts)   # pass previous hit to continue
print(hits, "cells match")
changed = wb.worksheets[0].replace("OldName", "NewName")
print(changed, "cells replaced")
```

## Hyperlinks

`Worksheet.hyperlinks.add(cell_name, total_rows, total_columns, address)` links a cell or
range to a URL, another cell (`"Sheet1!B10"`), or an external file. If the cell is empty,
the address is written as the cell's text; if it already has a value, that value shows as
plain text and you must style it to look like a link (see styles-and-formatting.md).

```python
import aspose.cells as gc

wb = gc.Workbook()
sheet = wb.worksheets[0]
sheet.hyperlinks.add("A1", 1, 1, "https://www.aspose.com")  # external URL
sheet.hyperlinks.add("A2", 1, 1, "Sheet1!B10")             # in-workbook target
sheet.hyperlinks.add("A3", 1, 1, "report.xlsx")            # external file
```

## Related
- styles-and-formatting.md - number/date display formats, aligning imported values, link styling.
- performance-and-large-files.md - streaming huge sheets, MemorySetting, LightCells import/export.
- formulas-and-calculation.md - reading computed results, calculate_formula.
- worksheets-rows-columns.md - inserting/deleting rows and columns, ranges, validation.
