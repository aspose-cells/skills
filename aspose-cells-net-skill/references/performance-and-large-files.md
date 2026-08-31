# Performance, Large Files, Threading

Read when a Workbook throws OutOfMemoryException or holds multi-GB RAM, bulk read/write is slow, CalculateFormula throws StackOverflowException, you use Aspose.Cells across threads, or a long load/save/convert/calc must be cancellable.

## Pitfalls

### OutOfMemoryException or multi-GB RAM on a large workbook

Symptom: the `Workbook` constructor or a bulk build throws `System.OutOfMemoryException`, or the process holds several GB for one file.
Cause: the default `MemorySetting.Normal` builds the full in-memory cell model and retains every cell object.
Fix (reading): pass `MemorySetting.MemoryPreference` through `LoadOptions`.

```csharp
var lo = new LoadOptions(LoadFormat.Xlsx) { MemorySetting = MemorySetting.MemoryPreference };
var wb = new Workbook("big.xlsx", lo);
// Then access cells sequentially: row by row, and cell by cell within each row.
```

Fix (building): set `WorkbookSettings.MemorySetting` before creating the sheets that hold the data. It becomes the default for worksheets created afterward and does NOT change worksheets that already exist (including `Worksheets[0]` from `new Workbook()`).

```csharp
var wb = new Workbook();
wb.Settings.MemorySetting = MemorySetting.MemoryPreference;   // default for sheets created next
var cells = wb.Worksheets[wb.Worksheets.Add()].Cells;        // this new sheet uses MemoryPreference
for (int r = 0; r < 100000; r++) cells[r, 0].PutValue(r);
wb.Save("big.xlsx");
```

Caveat: `MemoryPreference` trades speed for size. It slows random or repeated cell access and heavy insert/delete of cells and rows. Use it for read-mostly or sequential-write workloads; keep `Normal` when you edit heavily or touch cells out of order. It helps most when cells are numeric, boolean, or empty; for mostly-string or formula cells the memory cost is about the same as `Normal`.

### Files too big to hold in memory at all: LightCells streaming

Symptom: even `MemoryPreference` is not enough; the full model never fits.
Cause: every non-streaming path materializes cells into the workbook.
Fix: process cells one at a time and discard each. Save via `LightCellsDataProvider`, read via `LightCellsDataHandler`. Streaming an XLSX save commonly cuts memory by 50% or more versus the normal path.

Write path - implement the provider. Only the current cell is live; produce rows and columns in strictly ascending order, returning -1 to end.

```csharp
class MatrixProvider : LightCellsDataProvider
{
    int row = -1, col = -1;
    const int Rows = 100000, Cols = 30;
    public bool StartSheet(int sheetIndex) { return sheetIndex == 0; }  // true = stream this sheet
    public int NextRow() { row++; col = -1; return row < Rows ? row : -1; }  // -1 ends the sheet
    public void StartRow(Row r) { }                                    // set row-level props here
    public int NextCell() { col++; return col < Cols ? col : -1; }     // -1 ends the current row
    public void StartCell(Cell c) { c.PutValue(row * Cols + col); }    // fill the current cell
    public bool IsGatherString() { return false; }
}
```

```csharp
var wb = new Workbook();
var opts = new OoxmlSaveOptions(SaveFormat.Xlsx) { LightCellsDataProvider = new MatrixProvider() };
wb.Save("big.xlsx", opts);
```

Read path - implement the handler. Return `true` from each `Start*` and from `ProcessRow` to descend into that sheet, row, or cell; the `Cell` passed to `ProcessCell` is valid only inside that call.

```csharp
class CountingHandler : LightCellsDataHandler
{
    public int Total;
    public bool StartSheet(Worksheet sheet) { return true; }    // false skips the whole sheet
    public bool StartRow(int rowIndex) { return true; }         // false skips this row
    public bool ProcessRow(Row row) { return true; }            // true = visit this row's cells
    public bool StartCell(int columnIndex) { return true; }     // false skips this cell
    public bool ProcessCell(Cell cell) { Total++; return true; }
}
```

```csharp
var handler = new CountingHandler();
var lo = new LoadOptions(LoadFormat.Xlsx) { LightCellsDataHandler = handler };
var wb = new Workbook("big.xlsx", lo);   // cells stream through the handler during load
Console.WriteLine(handler.Total);
```

Constraints: access is strictly sequential (ascending rows, then ascending columns) with no random cell access, and features that need the full model are unavailable while streaming.

### Random corruption or exceptions under parallelism

Symptom: intermittent wrong values, null references, or corrupt output when several threads touch Aspose.Cells.
Cause: a `Workbook` and its object tree (worksheets, cells, styles) are NOT thread-safe.
Fix: give each thread or task its own `Workbook`; never share one instance across threads for writing. For concurrent read-only access to a single workbook, opt in explicitly:

```csharp
var wb = new Workbook("data.xlsx");
var cells = wb.Worksheets[0].Cells;
cells.MultiThreadReading = true;   // set before starting the reader threads
var t1 = new System.Threading.Thread(() => { _ = cells[0, 0].IntValue; });
var t2 = new System.Threading.Thread(() => { _ = cells[9, 0].IntValue; });
t1.Start(); t2.Start(); t1.Join(); t2.Join();
```

`MultiThreadReading` covers raw cell values only. Formatted or display string values (for example `StringValue`, `DisplayStringValue`) can still be wrong under threads, and it never makes concurrent writes safe.

### StackOverflowException inside CalculateFormula (IIS / ASP.NET)

Symptom: `Workbook.CalculateFormula` throws `System.StackOverflowException` on IIS or in an ASP.NET worker (it cannot be caught and kills the process).
Cause: IIS worker threads default to a small stack (about 256 KB); deep formula dependency chains recurse past it.
Fix: run the calculation on a dedicated thread with a large stack.

```csharp
var wb = new Workbook("deep-formulas.xlsx");
// IIS worker stack is ~256 KB; give the calc a 4 MB stack instead.
var t = new System.Threading.Thread(() => wb.CalculateFormula(), 4 * 1024 * 1024);
t.Start();
t.Join();
```

`CalculationOptions.CalcStackSize` (int, default 200) caps how many cells the engine calculates recursively; pass it through `CalculateFormula` when tuning engine recursion depth alongside the OS thread stack.

```csharp
var wb = new Workbook("deep-formulas.xlsx");
var opts = new CalculationOptions { CalcStackSize = 200 };   // recursion cap, default 200
wb.CalculateFormula(opts);
```

### Long load, save, convert, or calc that must time out

Symptom: a single load, save, PDF conversion, or `CalculateFormula` runs for minutes with no way to cancel.
Cause: these operations are CPU- and memory-intensive and run to completion by default.
Fix (load/save/convert): attach an `InterruptMonitor` and trip it from a watchdog. Use `Workbook.InterruptMonitor` for save and convert, and `LoadOptions.InterruptMonitor` for load.

```csharp
var wb = new Workbook("huge.xlsx");
var im = new InterruptMonitor();
wb.InterruptMonitor = im;   // for load, set LoadOptions.InterruptMonitor before constructing
var watchdog = new System.Threading.Thread(() =>
{
    System.Threading.Thread.Sleep(30000);
    im.Interrupt();         // aborts the running operation
});
watchdog.Start();
wb.Save("huge.pdf");        // throws a CellsException when interrupted; wrap in try/catch
```

Fix (formula calc): formula calculation uses a different hook. Derive from `AbstractCalculationMonitor`, call `Interrupt(msg)` from `BeforeCalculate`, and pass the instance through `CalculationOptions.CalculationMonitor`.

```csharp
class CalcInterrupter : AbstractCalculationMonitor
{
    public int Budget = 500000;   // stop after this many cells
    public override void BeforeCalculate(int sheetIndex, int rowIndex, int colIndex)
    {
        if (--Budget <= 0) Interrupt("calculation budget exceeded");   // aborts the whole CalculateFormula call
    }
}
```

```csharp
var wb = new Workbook("deep-formulas.xlsx");
var opts = new CalculationOptions { CalculationMonitor = new CalcInterrupter() };
wb.CalculateFormula(opts);   // returns early once the monitor interrupts
```

## Fast bulk write

Symptom: building a sheet with per-cell `PutValue` in a tight loop is slow.
Fix: import in blocks with `Cells.ImportData` / `ImportArray` instead of per-cell writes (see cell-values-and-data.md). During heavy bulk edits, suppress repeated recalculation and calculate once at the end (see formulas-and-calculation.md).

## Iteration and MaxDataRow

Cache the bounds before the loop: never call `Cells.MaxDataRow` / `MaxDataColumn` in a loop condition, since each call rescans the collection. Row and column enumeration and its performance notes are owned by cell-values-and-data.md.

## AutoFit cost

`AutoFitColumns` / `AutoFitRows` measure text for every affected cell and are expensive on large sheets; call them once after loading data, never inside a loop. See worksheets-rows-columns.md.

## Related
- cell-values-and-data.md - row enumeration and bulk import patterns.
- formulas-and-calculation.md - calculation chain and per-cell CalculationOptions.
- worksheets-rows-columns.md - AutoFit cost and structure operations.
- conversion-and-rendering.md - save options for long conversions.
