# Smart Markers and Reporting

Read this when generating reports from a pre-styled template workbook (markers like `&=Table.Field` or `&=$Variable` in the cells) using `WorkbookDesigner`, or when the marker cells are left blank or not replaced after `process()`.

Smart markers are tokens placed in a template workbook; `WorkbookDesigner.process()` fills them from data sources you register with `set_data_source`. They are the fastest way to produce a formatted report because the template already carries styles, layout, headers, and grouping.

## Pitfalls

### Marker cell left blank / not replaced

Symptom: after `process()` the cell still shows `&=Customer.City` or is empty.
Cause: the data source name or field does not match. The `&=Table.Field` prefix before the dot (`Customer`) must equal the name passed to `set_data_source`, and `Field` must be a real column/member; a mismatch silently leaves the cell empty. Variable markers `&=$Name` must match the name given to `set_data_source` exactly (case-sensitive, no spaces).
Fix: match the token to the registered name; confirm the field exists in the source. Use `&=$Name` for a single scalar value and `&=Source.Field` for a row-expanding table.

### Forgot process()

Symptom: the template is saved unchanged.
Cause: markers are only expanded by `process()`; constructing the designer or calling `calculate_formula` does not expand them.
Fix: always call `designer.process()` before `designer.workbook.save(...)`.

### Repeated rows collapsed into one

Symptom: a table marker produced only the first row.
Cause: the marker must be in a single cell and Aspose expands it downward by default; if the marker cell got overwritten by a `put_value` or the source has no rows, only one (blank) result appears.
Fix: keep the marker in its own cell, register a non-empty source, and let `process` expand. Add the `horizontal` modifier (`&=Table.Field(horizontal)`) only if you want it to expand across columns.

### Smart markers and formulas

Symptom: a total formula next to a marker block is wrong.
Cause: `process()` expands rows after formula setup; the formula in the marker row is copied but downstream totals may need recalculation.
Fix: call `designer.workbook.calculate_formula()` after `process()` so totals over the expanded range are correct (see formulas-and-calculation.md).

## Variable markers

`&=$Name` binds to a single scalar registered with `set_data_source("Name", value)`.

```python
import aspose.cells as gc

designer_book = gc.Workbook("template.xlsx")   # a cell contains &=$ReportDate
designer = gc.WorkbookDesigner(designer_book)
designer.set_data_source("ReportDate", "2026-08-21")
designer.process()
designer.workbook.save("out.xlsx", gc.SaveFormat.XLSX)
```

## Class / object markers from a DataTable

`&=Source.Field` expands one row per source record, copying the marker cell's style down. Build a `System.Data.DataTable` and register it by the `Source` name.

```python
import aspose.cells as gc

designer_book = gc.Workbook("template.xlsx")   # B2 contains &=Customer (array marker)
designer = gc.WorkbookDesigner(designer_book)

# Array-style markers accept a plain 2D Python list as the data source (no pythonnet
# needed). For field-style markers (&=Customer.City / &=Customer.Revenue) you need a
# System.Data.DataTable, which requires the pythonnet package (`import System` available):
#   import System.Data as sd
#   dt = sd.DataTable("Customer")
#   dt.Columns.Add("City", sd.String); dt.Columns.Add("Revenue", sd.Double)
#   dt.Rows.Add(["NYC", 5000.0]); dt.Rows.Add(["LA", 3000.0])
#   designer.set_data_source("Customer", dt)
designer.set_data_source("Customer", [["NYC", 5000.0], ["LA", 3000.0]])
designer.process()
designer.workbook.calculate_formula()       # recompute any totals over the expanded rows
designer.workbook.save("report.xlsx", gc.SaveFormat.XLSX)
```

To bind a list of plain Python objects instead of a DataTable, register each via `set_data_source` against an `ICellsDataTable` implementation; `System.Data.DataTable` is the lowest-friction option when the .NET `System.Data` assembly is available.

## Common marker modifiers

Append inside the parentheses after the field:

- `(group:normal)` group repeated values and repeat the group header.
- `(ascending)` / `(descending)` sort the source data.
- `(horizontal)` expand across columns instead of down rows.
- `(noadd)` do not add extra rows; only fill existing.
- `&==Table.Field` repeats a formula across the expanded range instead of a literal value.

Example: `&=Customer.Revenue(descending)` fills revenue sorted high-to-low, copying the cell style.

## Replacing a whole template block

Use `&=Table` (no `.Field`) to dump an entire record set starting at the marker cell, one column per source column.

```python
import aspose.cells as gc

designer_book = gc.Workbook("template.xlsx")   # A1 contains &=Customer (array marker)
designer = gc.WorkbookDesigner(designer_book)

# A 2D list is an acceptable data source for array-style markers. Field-style markers
# (&=Customer.Name) require a System.Data.DataTable - see the note in the example above.
designer.set_data_source("Customer", [["Acme", "West"], ["Globex", "East"]])
designer.process()
designer.workbook.save("out.xlsx", gc.SaveFormat.XLSX)
```

## Reading markers already in a workbook

There is no API to enumerate markers; they are plain cell text. To find template tokens, scan with `FindOptions` for `&=` (see cell-values-and-data.md for `Cells.find`).

## Related
- cell-values-and-data.md - `Cells.find` to locate marker tokens in a template.
- formulas-and-calculation.md - recalculating totals after `process()`.
- styles-and-formatting.md - markers copy the marker cell's style; design the template there.
- conversion-and-rendering.md - render the finished report to PDF/image.
