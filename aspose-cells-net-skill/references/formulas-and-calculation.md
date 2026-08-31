# Formulas and Calculation

Read this when formulas return null/0/stale values, calculation is slow, circular
references compute to zero, `#NAME?` appears from unsupported or custom functions, or you
need shared/array/dynamic-array formulas or defined names.

Aspose.Cells stores a formula plus its *last cached result* and never auto-evaluates.
Excel recalculates on open; Aspose does not. Any computed value you read is stale until you
call a calculate method.

## Pitfalls

### Formula set but the cell value is null, empty, or stale

Symptom: after `cell.Formula = ...` (or after loading a file your code edited) reading
`cell.Value` / `cell.DoubleValue` returns null, 0, or the previous value.
Cause: `Formula` and `SetFormula(...)` only store the expression - they do not evaluate it.
A loaded file carries whatever result was cached when it was last saved by Excel or Aspose.
Fix: call `workbook.CalculateFormula()` after (re)setting formulas and before reading
results. For one sheet use `worksheet.CalculateFormula(options, recursive)`; for one cell
use `cell.Calculate(options)`.

```csharp
var workbook = new Workbook("input.xlsx");
Cell cell = workbook.Worksheets[0].Cells["A3"];
cell.Formula = "=SUM(A1:A2)";     // stored, NOT evaluated
object stale = cell.Value;        // still the OLD cached result
workbook.CalculateFormula();      // evaluate every formula in the book
object fresh = cell.Value;        // now correct
```

Direct calculation - evaluate an expression without writing it into a cell.
`Worksheet.CalculateFormula(string)` returns the result:

```csharp
var workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];
sheet.Cells["A1"].PutValue(20);
sheet.Cells["A2"].PutValue(30);
object sum = sheet.CalculateFormula("=SUM(A1:A2)");  // returns 50, no cell used
```

### CalculateFormula is slow or repeated recalcs are expensive

Symptom: a full `CalculateFormula()` takes seconds; recalculating after a small edit repeats
the whole cost.
Cause: by default no calculation chain is retained, so every call re-evaluates everything,
and `Cell.Calculate` with the default `Recursive = true` re-walks all dependents each call.
Fix (repeated partial recalcs): enable the chain once. The first pass costs more because it
builds the chain; later passes recompute only what changed. Leave it off for a single calc.

```csharp
var workbook = new Workbook("big.xlsx");
workbook.Settings.FormulaSettings.EnableCalculationChain = true;
workbook.CalculateFormula();                      // first pass builds the chain
workbook.Worksheets[0].Cells["A1"].PutValue(99);
workbook.CalculateFormula();                      // later passes recalc only what changed
```

Fix (many single-cell calcs in a loop): set `Recursive = false` so dependents are computed
once instead of re-walked on every call. Set `IgnoreError = true` so an unsupported function
or dead external link records a cell error value instead of throwing.

```csharp
var workbook = new Workbook("big.xlsx");
var opts = new CalculationOptions { Recursive = false, IgnoreError = true };
Cell cell = workbook.Worksheets[0].Cells["B2"];
cell.Calculate(opts);   // recompute just B2; dependents not re-walked each call
```

Avoid volatile functions (`NOW`, `TODAY`, `RAND`, `RANDBETWEEN`, `OFFSET`, `INDIRECT`,
`INFO`) where you can: they recompute on every pass and force their dependents to recompute,
defeating the calculation chain.

### Circular references compute to 0

Symptom: intentionally circular formulas (iterative models) all read 0.
Cause: iterative calculation is off by default, so a cycle cannot converge.
Fix: enable and bound iterative calculation, then calculate.

```csharp
var workbook = new Workbook("circular.xlsx");
FormulaSettings fs = workbook.Settings.FormulaSettings;
fs.EnableIterativeCalculation = true;   // off by default -> cycles resolve to 0
fs.MaxIteration = 100;
fs.MaxChange = 0.001;
workbook.CalculateFormula();
```

To *detect* (rather than resolve) cycles, assign a subclass of `AbstractCalculationMonitor`
to `CalculationOptions.CalculationMonitor` and collect the `CalculationCell` objects passed
to its `OnCircular` override.

### #NAME? from unsupported or custom (UDF) functions

Symptom: a cell shows `#NAME?` (or a blank/stale value) after calculation.
Cause: the function name is not a built-in Aspose supports, or it is your own UDF the engine
has no definition for. Aspose supports most Excel math, text, logical, date, statistical,
lookup, database, and financial functions; unknown names do not resolve on their own.
Fix: supply a custom engine. Subclass `AbstractCalculationEngine`, override
`Calculate(CalculationData)`, and write the result to `data.CalculatedValue`. The engine is
consulted for names Aspose does not recognize, so your UDF resolves instead of `#NAME?`.

```csharp
public class CustomFunctionEngine : AbstractCalculationEngine
{
    public override void Calculate(CalculationData data)
    {
        if (string.Equals(data.FunctionName, "MYADD", StringComparison.OrdinalIgnoreCase))
        {
            double a = Convert.ToDouble(data.GetParamValue(0));
            double b = Convert.ToDouble(data.GetParamValue(1));
            data.CalculatedValue = a + b;
        }
    }
}
```

```csharp
var workbook = new Workbook();
Cell cell = workbook.Worksheets[0].Cells["A1"];
cell.Formula = "=MYADD(2,3)";     // unknown function -> would be #NAME?
var opts = new CalculationOptions { CustomEngine = new CustomFunctionEngine() };
workbook.CalculateFormula(opts);
double result = cell.DoubleValue; // 5
```

`CalculationData` exposes `FunctionName`, `ParamCount`, `GetParamValue(int)`,
`GetParamText(int)`, `CellRow` / `CellColumn`, and `Worksheet`. To also override a built-in
function (for example force `TODAY` to a fixed date), override `ProcessBuiltInFunctions` to
return `true`. The old `ICustomFunction` interface is obsolete - use
`AbstractCalculationEngine`.

## Calculation mode and calculate-on-open

`FormulaSettings.CalculationMode` (`CalcModeType.Automatic`, `AutomaticExceptTable`,
`Manual`) and `CalculateOnOpen` are flags written into the saved file; they control what
*Excel* does when it opens the output. They do not make Aspose calculate at runtime - you
must still call `CalculateFormula` yourself, whatever the mode.

```csharp
var workbook = new Workbook("report.xlsx");
FormulaSettings fs = workbook.Settings.FormulaSettings;
fs.CalculationMode = CalcModeType.Manual;   // saved flag; affects Excel-on-open, not Aspose
fs.CalculateOnOpen = false;
workbook.Save("report.xlsx");
```

## Writing formulas

Set `Cell.Formula` (A1 style) or `Cell.R1C1Formula`; always begin with `=` and delimit
arguments with commas regardless of locale.

```csharp
Cells cells = new Workbook().Worksheets[0].Cells;
cells["A11"].Formula = "=SUM(A1:A10)";
cells["B11"].R1C1Formula = "=SUM(R[-10]C:R[-1]C)";   // references B1:B10
```

Shared and array formulas fill a block in one call. `SetSharedFormula(text, rows, cols)`
auto-shifts relative references down the block (cheaper to parse than a per-row formula).
`SetArrayFormula(text, rows, cols)` enters one CSE array formula spilling over the block.

```csharp
Cells cells = new Workbook().Worksheets[0].Cells;
cells["B2"].SetSharedFormula("=A2*0.09", 100, 1);   // 100 rows x 1 col, refs auto-shift
cells["D1"].SetArrayFormula("=A1:A3*2", 3, 1);      // spills over 3 rows x 1 col
```

Dynamic-array (spill) formulas: set with `SetDynamicArrayFormula(text, options, calculate)`,
then call `Workbook.RefreshDynamicArrayFormulas(true)` to size and fill the spilled range.

```csharp
var workbook = new Workbook("data.xlsx");
Cell cell = workbook.Worksheets[0].Cells["D1"];
cell.SetDynamicArrayFormula("=SORT(A1:A10)", new FormulaParseOptions(), true);
workbook.RefreshDynamicArrayFormulas(true);
```

## Defined names (named ranges)

`Worksheets.Names.Add(name)` returns an index; fetch the `Name` and set `RefersTo` to an
absolute reference (prefix a sheet name to scope it). Formulas can then use the name.

```csharp
var workbook = new Workbook();
int idx = workbook.Worksheets.Names.Add("TaxRate");
Name name = workbook.Worksheets.Names[idx];
name.RefersTo = "=Sheet1!$A$1";
Worksheet sheet = workbook.Worksheets[0];
sheet.Cells["A1"].PutValue(0.09);
sheet.Cells["B1"].Formula = "=100*TaxRate";
workbook.CalculateFormula();
```

## Interrupting a long calculation

To cancel a calculation that runs too long (via `AbstractCalculationMonitor.Interrupt(msg)`),
and to fix `StackOverflowException` under IIS or other limited-stack hosts using
`CalculationOptions.CalcStackSize`, see `performance-and-large-files.md`.

## Related
- charts.md - call chart.Calculate() before rendering a chart image.
- pivot-tables.md - refresh the cache and CalculateData() to populate pivot values.
- performance-and-large-files.md - InterruptMonitor, CalcStackSize, threading.
- cell-values-and-data.md - reading typed results back (Value, DoubleValue, StringValue).
- styles-and-formatting.md - display a computed number or date with a format.
