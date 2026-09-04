# Formulas and Calculation

Read this when formulas return null/0/stale values, calculation is slow, circular
references compute to zero, `#NAME?` appears from unsupported or custom functions, or you
need shared/array/dynamic-array formulas or defined names.

Aspose.Cells stores a formula plus its *last cached result* and never auto-evaluates.
Excel recalculates on open; Aspose does not. Any computed value you read is stale until you
call a calculate method.

## Pitfalls

### Formula set but the cell value is null, empty, or stale

Symptom: after `cell.formula = ...` (or after loading a file your code edited) reading
`cell.value` / `cell.double_value` returns null, 0, or the previous value.
Cause: `formula` only stores the expression - it does not evaluate it.
A loaded file carries whatever result was cached when it was last saved by Excel or Aspose.
Fix: call `workbook.calculate_formula()` after (re)setting formulas and before reading
results. For one sheet use `worksheet.calculate_formula(options, recursive)`; for one cell
use `cell.calculate(options)`.

```python
import aspose.cells as gc

workbook = gc.Workbook("input.xlsx")
cell = workbook.worksheets[0].cells.get("A3")
cell.formula = "=SUM(A1:A2)"     # stored, NOT evaluated
stale = cell.value              # still the OLD cached result
workbook.calculate_formula()     # evaluate every formula in the book
fresh = cell.value              # now correct
```

Direct calculation - evaluate an expression without writing it into a cell.
`Worksheet.calculate_formula(string)` returns the result:

```python
import aspose.cells as gc

workbook = gc.Workbook()
sheet = workbook.worksheets[0]
sheet.cells.get("A1").put_value(20)
sheet.cells.get("A2").put_value(30)
total = sheet.calculate_formula("=SUM(A1:A2)")  # returns 50, no cell used
```

### CalculateFormula is slow or repeated recalcs are expensive

Symptom: a full `calculate_formula()` takes seconds; recalculating after a small edit repeats
the whole cost.
Cause: by default no calculation chain is retained, so every call re-evaluates everything.
Fix (repeated partial recalcs): enable the chain once. The first pass costs more because it
builds the chain; later passes recompute only what changed. Leave it off for a single calc.

```python
import aspose.cells as gc

workbook = gc.Workbook("big.xlsx")
workbook.settings.formula_settings.enable_calculation_chain = True
workbook.calculate_formula()                      # first pass builds the chain
workbook.worksheets[0].cells.get("A1").put_value(99)
workbook.calculate_formula()                      # later passes recalc only what changed
```

Fix (many single-cell calcs in a loop): set `recursive = False` so dependents are computed
once instead of re-walked on every call. Set `ignore_error = True` so an unsupported function
or dead external link records a cell error value instead of throwing.

```python
import aspose.cells as gc

workbook = gc.Workbook("big.xlsx")
opts = gc.CalculationOptions()
opts.recursive = False
opts.ignore_error = True
cell = workbook.worksheets[0].cells.get("B2")
cell.calculate(opts)   # recompute just B2; dependents not re-walked each call
```
Avoid volatile functions (`NOW`, `TODAY`, `RAND`, `RANDBETWEEN`, `OFFSET`, `INDIRECT`,
`INFO`) where you can: they recompute on every pass and force their dependents to recompute.

### Controlling external link updates on load

Symptom: opening a workbook with external links either blocks waiting for user input, or
automatically updates links you did not intend to refresh.
Cause: `Workbook.settings.update_links_type` controls whether external links are updated
when the workbook opens. The default is `USER_SET` (prompt the user).
Fix: set the property before or after loading; no `LoadOptions` parameter is needed.

```python
import aspose.cells as gc

wb = gc.Workbook("file_with_links.xlsx")
wb.settings.update_links_type = gc.UpdateLinksType.NEVER   # never update external links
# wb.settings.update_links_type = gc.UpdateLinksType.ALWAYS  # always update
# wb.settings.update_links_type = gc.UpdateLinksType.USER_SET  # prompt user (default)
```

`UpdateLinksType` enum values: `USER_SET` (0), `NEVER` (1), `ALWAYS` (2).

### Circular references compute to 0

Symptom: intentionally circular formulas (iterative models) all read 0.
Cause: iterative calculation is off by default, so a cycle cannot converge.
Fix: enable and bound iterative calculation, then calculate.

```python
import aspose.cells as gc

workbook = gc.Workbook("circular.xlsx")
fs = workbook.settings.formula_settings
fs.enable_iterative_calculation = True   # off by default -> cycles resolve to 0
fs.max_iteration = 100
fs.max_change = 0.001
workbook.calculate_formula()
```

### #NAME? from unsupported or custom (UDF) functions

Symptom: a cell shows `#NAME?` (or a blank/stale value) after calculation.
Cause: the function name is not a built-in Aspose supports, or it is your own UDF the engine
has no definition for. Aspose supports most Excel math, text, logical, date, statistical,
lookup, database, and financial functions; unknown names do not resolve on their own.
Fix (when your wheel supports it): supply a custom engine by subclassing
`AbstractCalculationEngine`, overriding `calculate(data)`, and writing the result to
`data.calculated_value`. The engine is consulted for names Aspose does not recognize, so
your UDF resolves instead of `#NAME?`. Note: in some `aspose-cells-python` wheels subclassing
`AbstractCalculationEngine` raises `TypeError: not an acceptable base type`; in that build
custom UDFs are unavailable and the name stays `#NAME?`.

```python
import aspose.cells as gc

workbook = gc.Workbook()
cell = workbook.worksheets[0].cells.get("A1")
cell.formula = "=SUM(1,2,3)"
opts = gc.CalculationOptions()
opts.calc_stack_size = 200       # engine recursion cap (default 200) for deep chains
workbook.calculate_formula(opts)
print(cell.double_value)         # 6
```

`CalculationData` exposes `function_name`, `param_count`, `get_param_value(index)`,
`cell_row` / `cell_column`, and `worksheet`. To also override a built-in
function (for example force `TODAY` to a fixed date), override `process_builtin_functions`
to return `True`. The old `ICustomFunction` interface is obsolete - use
`AbstractCalculationEngine` (subject to the subclassing caveat above).

## Calculation mode and calculate-on-open

`FormulaSettings.calculation_mode` (`CalcModeType.AUTOMATIC`, `AUTOMATIC_EXCEPT_TABLE`,
`MANUAL`) and `calculate_on_open` are flags written into the saved file; they control what
*Excel* does when it opens the output. They do not make Aspose calculate at runtime - you
must still call `calculate_formula` yourself, whatever the mode.

```python
import aspose.cells as gc

workbook = gc.Workbook("report.xlsx")
fs = workbook.settings.formula_settings
fs.calculation_mode = gc.CalcModeType.MANUAL   # saved flag; affects Excel-on-open, not Aspose
fs.calculate_on_open = False
workbook.save("report.xlsx")
```

## Writing formulas

Set `Cell.formula` (A1 style) or `Cell.r1_c1_formula`; always begin with `=` and delimit
arguments with commas regardless of locale.

```python
import aspose.cells as gc

cells = gc.Workbook().worksheets[0].cells
cells.get("A11").formula = "=SUM(A1:A10)"
cells.get("B11").r1c1_formula = "=SUM(R[-10]C:R[-1]C)"   # references B1:B10
```

Shared and array formulas fill a block in one call. `set_shared_formula(text, rows, cols)`
auto-shifts relative references down the block. `set_array_formula(text, rows, cols)` enters
one CSE array formula spilling over the block.

```python
import aspose.cells as gc

cells = gc.Workbook().worksheets[0].cells
cells.get("B2").set_shared_formula("=A2*0.09", 100, 1)   # 100 rows x 1 col, refs auto-shift
cells.get("D1").set_array_formula("=A1:A3*2", 3, 1)      # spills over 3 rows x 1 col
```

Dynamic-array (spill) formulas: set with `set_dynamic_array_formula(text, options, calculate)`,
then call `Workbook.refresh_dynamic_array_formulas(True)` to size and fill the spilled range.

```python
import aspose.cells as gc

workbook = gc.Workbook("data.xlsx")
cell = workbook.worksheets[0].cells.get("D1")
cell.set_dynamic_array_formula("=SORT(A1:A10)", gc.FormulaParseOptions(), True)
workbook.refresh_dynamic_array_formulas(True)
```

## Defined names (named ranges)

`Worksheets.names.add(name)` returns an index; fetch the `Name` and set `refers_to` to an
absolute reference (prefix a sheet name to scope it). Formulas can then use the name.

```python
import aspose.cells as gc

workbook = gc.Workbook()
idx = workbook.worksheets.names.add("TaxRate")
name = workbook.worksheets.names[idx]
name.refers_to = "=Sheet1!$A$1"
sheet = workbook.worksheets[0]
sheet.cells.get("A1").put_value(0.09)
sheet.cells.get("B1").formula = "=100*TaxRate"
workbook.calculate_formula()
```

## Interrupting a long calculation

To cancel a calculation that runs too long (via `AbstractCalculationMonitor.interrupt(msg)`),
and to fix `RecursionError` under limited-stack hosts using
`CalculationOptions.calc_stack_size`, see `performance-and-large-files.md`.

## Related
- charts.md - call chart.calculate() before rendering a chart image.
- pivot-tables.md - refresh the cache and calculate_data() to populate pivot values.
- performance-and-large-files.md - InterruptMonitor, CalcStackSize, threading.
- cell-values-and-data.md - reading typed results back (value, double_value, string_value).
- styles-and-formatting.md - display a computed number or date with a format.
