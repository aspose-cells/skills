# Performance, Large Files, Threading

Read when a Workbook throws OutOfMemoryError or holds multi-GB RAM, bulk read/write is slow, calculate_formula throws RecursionError, you use Aspose.Cells across threads, or a long load/save/convert/calc must be cancellable.

## Pitfalls

### OutOfMemoryError or multi-GB RAM on a large workbook

Symptom: the `Workbook` constructor or a bulk build raises `OutOfMemoryError`, or the process holds several GB for one file.
Cause: the default `MemorySetting.NORMAL` builds the full in-memory cell model and retains every cell object.
Fix (reading): pass `MemorySetting.MEMORY_PREFERENCE` through `LoadOptions`.

```python
import aspose.cells as gc

lo = gc.LoadOptions(gc.LoadFormat.XLSX)
lo.memory_setting = gc.MemorySetting.MEMORY_PREFERENCE
wb = gc.Workbook("big.xlsx", lo)
# Then access cells sequentially: row by row, and cell by cell within each row.
```

Fix (building): set `Workbook.settings.memory_setting` before creating the sheets that hold the data. It becomes the default for worksheets created afterward and does NOT change worksheets that already exist (including `Worksheets[0]` from `new Workbook()`).

```python
import aspose.cells as gc

wb = gc.Workbook()
wb.settings.memory_setting = gc.MemorySetting.MEMORY_PREFERENCE   # default for sheets created next
cells = wb.worksheets[wb.worksheets.add()].cells        # this new sheet uses MEMORY_PREFERENCE
for r in range(100000):
    cells.get(r, 0).put_value(r)
wb.save("big.xlsx")
```

Caveat: `MEMORY_PREFERENCE` trades speed for size. It slows random or repeated cell access and heavy insert/delete of cells and rows. Use it for read-mostly or sequential-write workloads; keep `NORMAL` when you edit heavily or touch cells out of order. It helps most when cells are numeric, boolean, or empty; for mostly-string or formula cells the memory cost is about the same as `NORMAL`.

### Files too big to hold in memory at all: LightCells streaming

Symptom: even `MEMORY_PREFERENCE` is not enough; the full model never fits.
Cause: every non-streaming path materializes cells into the workbook.
Fix: process cells one at a time and discard each. Save via `LightCellsDataProvider`, read via `LightCellsDataHandler`. Streaming an XLSX save commonly cuts memory by 50% or more versus the normal path.

Write path - implement the provider. Only the current cell is live; produce rows and columns in strictly ascending order, returning -1 to end.

```python
import aspose.cells as gc

class MatrixProvider(gc.LightCellsDataProvider):
    def __init__(self):
        self.row = -1
        self.col = -1
        self.rows = 100000
        self.cols = 30

    def start_sheet(self, sheet_index):
        return sheet_index == 0   # True = stream this sheet

    def next_row(self):
        self.row += 1
        self.col = -1
        return self.row if self.row < self.rows else -1

    def start_row(self, row):
        pass

    def next_cell(self):
        self.col += 1
        return self.col if self.col < self.cols else -1

    def start_cell(self, cell):
        cell.put_value(self.row * self.cols + self.col)

    def is_gather_string(self):
        return False

wb = gc.Workbook()
opts = gc.OoxmlSaveOptions(gc.SaveFormat.XLSX)
opts.light_cells_data_provider = MatrixProvider()
wb.save("big.xlsx", opts)
```

Read path - implement the handler. Return `True` from each `start_*` and from `process_row` to descend into that sheet, row, or cell; the `Cell` passed to `process_cell` is valid only inside that call.

```python
import aspose.cells as gc

class CountingHandler(gc.LightCellsDataHandler):
    def __init__(self):
        self.total = 0

    def start_sheet(self, sheet):
        return True     # False skips the whole sheet

    def start_row(self, row_index):
        return True     # False skips this row

    def process_row(self, row):
        return True     # True = visit this row's cells

    def start_cell(self, column_index):
        return True     # False skips this cell

    def process_cell(self, cell):
        self.total += 1
        return True

handler = CountingHandler()
lo = gc.LoadOptions(gc.LoadFormat.XLSX)
lo.light_cells_data_handler = handler
wb = gc.Workbook("big.xlsx", lo)   # cells stream through the handler during load
print(handler.total)
```

Constraints: access is strictly sequential (ascending rows, then ascending columns) with no random cell access, and features that need the full model are unavailable while streaming.

### Random corruption or exceptions under parallelism

Symptom: intermittent wrong values, null references, or corrupt output when several threads touch Aspose.Cells.
Cause: a `Workbook` and its object tree (worksheets, cells, styles) are NOT thread-safe.
Fix: give each thread or task its own `Workbook`; never share one instance across threads for writing. For concurrent read-only access to a single workbook, opt in explicitly:

```python
import aspose.cells as gc
import threading

wb = gc.Workbook("data.xlsx")
cells = wb.worksheets[0].cells
cells.multi_thread_reading = True   # set before starting the reader threads
t1 = threading.Thread(target=lambda: cells.get(0, 0).int_value)
t2 = threading.Thread(target=lambda: cells.get(9, 0).int_value)
t1.start(); t2.start(); t1.join(); t2.join()
```

`multi_thread_reading` covers raw cell values only. Formatted or display string values (for example `string_value`, `display_string_value`) can still be wrong under threads, and it never makes concurrent writes safe.

### RecursionError inside calculate_formula (limited-stack hosts)

Symptom: `Workbook.calculate_formula` raises `RecursionError` on a host with a small stack, or when formula dependency chains are very deep.
Cause: deep formula dependency chains recurse past the available stack.
Fix: run the calculation on a dedicated thread with a large stack, or cap the engine recursion depth.

```python
import aspose.cells as gc
import threading

wb = gc.Workbook("deep-formulas.xlsx")
# Give the calc a 4 MB stack instead of the default.
threading.stack_size(4 * 1024 * 1024)
t = threading.Thread(target=wb.calculate_formula)
t.start(); t.join()
```

`CalculationOptions.calc_stack_size` (int, default 200) caps how many cells the engine calculates recursively; pass it through `calculate_formula` when tuning engine recursion depth alongside the OS thread stack.

```python
import aspose.cells as gc

wb = gc.Workbook("deep-formulas.xlsx")
opts = gc.CalculationOptions()
opts.calc_stack_size = 200   # recursion cap, default 200
wb.calculate_formula(opts)
```

### Long load, save, convert, or calc that must time out

Symptom: a single load, save, PDF conversion, or `calculate_formula` runs for minutes with no way to cancel.
Cause: these operations are CPU- and memory-intensive and run to completion by default.
Fix (load/save/convert): attach an `InterruptMonitor` and trip it from a watchdog. Use `Workbook.interrupt_monitor` for save and convert, and `LoadOptions.interrupt_monitor` for load.

```python
import aspose.cells as gc
import threading

wb = gc.Workbook("huge.xlsx")
im = gc.InterruptMonitor()
wb.interrupt_monitor = im   # for load, set LoadOptions.interrupt_monitor before constructing
def watchdog():
    threading.Event().wait(30)
    im.interrupt()          # aborts the running operation
threading.Thread(target=watchdog).start()
try:
    wb.save("huge.pdf")     # raises a CellsException when interrupted
except Exception as e:
    print("interrupted:", e)
```

Fix (formula calc): formula calculation uses a different hook. Derive from `AbstractCalculationMonitor`, call `interrupt(msg)` from `before_calculate`, and pass the instance through `CalculationOptions.calculation_monitor`.

```python
import aspose.cells as gc

# Tune the engine without a custom monitor:
wb = gc.Workbook()
opts = gc.CalculationOptions()
opts.calc_stack_size = 1000          # raise the recursion cap for deeply nested formulas
opts.ignore_error = True             # don't stop on per-cell calculation errors
wb.calculate_formula(opts)

# For progress reporting / interruption, Aspose's .NET pattern is to subclass
# AbstractCalculationMonitor and set opts.calculation_monitor. In some wheels that
# base is not subclassable (TypeError: not an acceptable base type); confirm your
# build supports it before relying on a custom monitor.
```

## Fast bulk write

Symptom: building a sheet with per-cell `put_value` in a tight loop is slow.
Fix: import in blocks with `Cells.import_array` / `import_object_array` instead of per-cell writes (see cell-values-and-data.md). During heavy bulk edits, suppress repeated recalculation and calculate once at the end (see formulas-and-calculation.md).

## Iteration and MaxDataRow

Cache the bounds before the loop: never call `Cells.max_data_row` / `max_data_column` in a loop condition, since each call rescans the collection. Row and column enumeration and its performance notes are owned by cell-values-and-data.md.

## AutoFit cost

`auto_fit_columns` / `auto_fit_rows` measure text for every affected cell and are expensive on large sheets; call them once after loading data, never inside a loop. See worksheets-rows-columns.md.

## Related
- cell-values-and-data.md - row enumeration and bulk import patterns.
- formulas-and-calculation.md - calculation chain and per-cell CalculationOptions.
- worksheets-rows-columns.md - AutoFit cost and structure operations.
- conversion-and-rendering.md - save options for long conversions.
