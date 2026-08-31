# Cell Values and Data IO

Read this when setting or reading cell values, getting unexpected runtime types or
formatted strings on read, importing/exporting DataTable/arrays/objects, iterating
cells, doing find/replace, or adding hyperlinks.

All row/column indices are 0-based (row 0 = Excel row 1). Import and `PutValue` write
VALUES ONLY - number/date display formats are a separate step; see
styles-and-formatting.md. Get a cell with `cells["A1"]` or `cells[row, col]`.

## Pitfalls

### Cell.Value returns an unpredictable boxed type

Symptom: `(double)cell.Value` throws `InvalidCastException`; a date cell yields a
number; an integer sometimes boxes as `int`, sometimes as `double`.
Cause: `Cell.Value` is `object`. Its runtime type is one of `null`, `bool`, `DateTime`,
`double`, `int`, `string` - and Aspose does not guarantee a stable type (an integer may
box as `int` or `double` across calls/versions). Dates are stored as OADate serials, so
a date cell's `Value` may box as `DateTime` OR surface as a `double`.
Fix: never cast `Cell.Value` directly. Branch on `Cell.Type` (a `CellValueType`) and
read through the typed accessors, which coerce reliably.

| Accessor | Returns | Use when |
|---|---|---|
| `Cell.Value` | `object` (unstable type) | you will type-check it yourself |
| `Cell.DoubleValue` / `FloatValue` / `IntValue` | number | `Type == IsNumeric` |
| `Cell.DateTimeValue` | `DateTime` | `Type == IsDateTime` |
| `Cell.BoolValue` | `bool` | `Type == IsBool` |
| `Cell.StringValue` | formatted `string` | any type |

`CellValueType` members: `IsNull`, `IsNumeric`, `IsDateTime`, `IsString`, `IsBool`,
`IsError`, `IsUnknown`. `DoubleValue`/`FloatValue` throw on an empty (`IsNull`) cell -
guard first.

```csharp
Workbook wb = new Workbook("input.xlsx");
Cell cell = wb.Worksheets[0].Cells["A1"];
switch (cell.Type)
{
    case CellValueType.IsNumeric:  Console.WriteLine(cell.DoubleValue);   break;
    case CellValueType.IsDateTime: Console.WriteLine(cell.DateTimeValue); break;
    case CellValueType.IsBool:     Console.WriteLine(cell.BoolValue);     break;
    case CellValueType.IsString:   Console.WriteLine(cell.StringValue);   break;
    case CellValueType.IsError:    Console.WriteLine(cell.StringValue);   break; // "#DIV/0!"
    case CellValueType.IsNull:     break;                                        // empty
}
```

### Read string differs from what Excel shows

Symptom: a cell displays `0.01` in Excel but your read returns `0.012345`, or vice versa.
Cause: three different strings exist for one cell. `StringValue` applies the cell's own
number format (identical to Excel "copy as text"). `DisplayStringValue` applies the
cell's display style - exactly what Excel renders (reflects conditional formatting).
The underlying unformatted value is neither.
Fix: pick the accessor that matches intent. For explicit control call
`GetStringValue(CellValueFormatStrategy)`: `None` = unformatted, `CellStyle` = cell's
format, `DisplayStyle` = as displayed.

```csharp
Workbook wb = new Workbook("input.xlsx");
Cell cell = wb.Worksheets[0].Cells["A1"];              // 0.012345, number format "0.00"
Console.WriteLine(cell.StringValue);                  // "0.01"     - cell number format
Console.WriteLine(cell.DisplayStringValue);           // "0.01"     - what Excel shows
Console.WriteLine(cell.GetStringValue(CellValueFormatStrategy.None)); // "0.012345" - raw
```

### Text numbers stay text and break SUM

Symptom: imported/typed numbers are left-aligned, `SUM` ignores them, sorting is
lexical. Cause: a string like `"00123"` written as text stays text. `PutValue(string)`
and `Cell.Value = "..."` do NOT parse. Fix: parse on write with the `PutValue(string,
bool isConverted)` overload (`true` converts to number/date/bool), or fix a whole sheet
after load with `Cells.ConvertStringToNumericValue()`.

```csharp
Workbook wb = new Workbook();
Cells cells = wb.Worksheets[0].Cells;
cells["A1"].PutValue("00123");        // stays text -> left-aligned, excluded from SUM
cells["A2"].PutValue("00123", true);  // parsed to numeric 123
cells.ConvertStringToNumericValue();  // whole sheet: convert every text-number in place
```

### Cells.ImportDataTable is obsolete

Symptom: build warning on `ImportDataTable`; options you need are missing.
Cause: the boolean-argument `ImportDataTable`/`ImportDataView` overloads are deprecated.
Fix: use `Cells.ImportData(DataTable, int firstRow, int firstColumn, ImportTableOptions)`
(overloads also accept `DataView`, `ICellsDataTable`, `IDataReader`). Configure behavior
through `ImportTableOptions` instead of positional flags.

Key `ImportTableOptions` members (all verified): `IsFieldNameShown` (write column names
as the first row), `ConvertNumericData` (parse numeric-looking strings), `DateFormat`
and `NumberFormats` (format strings applied to imported dates/numbers), `ColumnIndexes`
(0-based columns to import - the DataColumn selection path), `CheckMergedCells`,
`InsertRows`, `ShiftFirstRowDown`, `TotalColumns`, `IsHtmlString`.

```csharp
DataTable dt = new DataTable("Products");
dt.Columns.Add("Name", typeof(string));
dt.Columns.Add("Price", typeof(double));
dt.Rows.Add("Pen", 1.5);
dt.Rows.Add("Cup", 3.0);
Workbook wb = new Workbook();
ImportTableOptions opts = new ImportTableOptions();
opts.IsFieldNameShown = true;    // header row
opts.ConvertNumericData = true;  // parse numeric strings
wb.Worksheets[0].Cells.ImportData(dt, 0, 0, opts);
```

Import writes values only; apply date/number display formats afterward (see
styles-and-formatting.md).

### Scanning with Cells[r, c] is slow and instantiates empty cells

Symptom: iterating a large or sparse sheet is slow or throws `OutOfMemoryException`.
Cause: `cells[r, c]` and `cells["A1"]` CREATE a blank `Cell` when none exists, so a full
`r x c` scan instantiates the entire matrix. `MaxDataRow`/`MaxDataColumn` also recompute
statistics on every access.
Fix: enumerate only the cells that exist with the `Rows` -> `Row` -> `Cell` iterators.
Do not add or delete cells while enumerating (the enumerator may skip or repeat). If you
must loop by index, cache `MaxDataRow`/`MaxDataColumn` once - both return **-1** on an
empty sheet, so `for (r = 0; r <= maxRow; r++)` correctly runs zero times.

```csharp
Workbook wb = new Workbook("input.xlsx");
Cells cells = wb.Worksheets[0].Cells;
foreach (Row row in cells.Rows)          // only rows that exist
{
    foreach (Cell cell in row)           // only instantiated cells in that row
    {
        if (cell.Type == CellValueType.IsNull) continue;
        Console.WriteLine(cell.Row + "," + cell.Column + " = " + cell.StringValue);
    }
}
```

`Cells`, `Row`, and `Range` each expose `GetEnumerator()` yielding `Cell`; `Row` also
gives `FirstCell`, `LastCell`, `LastDataCell`, and `Index`.

## Setting values

`Cell.PutValue` has typed overloads - `PutValue(double)`, `PutValue(int)`,
`PutValue(bool)`, `PutValue(DateTime)`, `PutValue(string)`, plus `PutValue(object)` and
the parsing `PutValue(string, bool)`. Setting `Cell.Value = x` is equivalent to the
matching typed overload. For bulk writes, populate first by row then by column - it is
markedly faster than column-major order.

```csharp
Workbook wb = new Workbook();
Cells cells = wb.Worksheets[0].Cells;
cells["A1"].PutValue("Total");     // string
cells["B1"].PutValue(42.5);        // double
cells["C1"].PutValue(true);        // bool
cells["D1"].PutValue(DateTime.Now);// DateTime - still needs a date format to display right
cells["E1"].Value = 7;             // equivalent to PutValue(7)
```

## Importing arrays and objects

`ImportArray` takes a 1D `string[]`, `int[]`, or `double[]` with
`(array, firstRow, firstColumn, isVertical)`. `ImportTwoDimensionArray(object[,],
firstRow, firstColumn)` lays a rectangular block. For a list of business objects,
`ImportCustomObjects(ICollection, string[] propertyNames, bool isPropertyNameShown, int
firstRow, int firstColumn, int rowNumber, bool insertRows, string dateFormatString, bool
convertStringToNumber)` maps named properties to columns.

```csharp
Workbook wb = new Workbook();
Cells cells = wb.Worksheets[0].Cells;
string[] header = { "Name", "Qty" };
cells.ImportArray(header, 0, 0, false);      // false = horizontal: A1, B1
double[] prices = { 1.5, 3.0, 2.25 };
cells.ImportArray(prices, 1, 0, true);       // true = vertical: A2, A3, A4
object[,] grid = { { "Pen", 12 }, { "Cup", 5 } };
cells.ImportTwoDimensionArray(grid, 1, 1);   // 2x2 block anchored at B2
```

```csharp
public class Product
{
    public string Name { get; set; }
    public double Price { get; set; }
}
```

```csharp
List<Product> items = new List<Product>
{
    new Product { Name = "Pen", Price = 1.5 },
    new Product { Name = "Cup", Price = 3.0 }
};
Workbook wb = new Workbook();
string[] props = { "Name", "Price" };
wb.Worksheets[0].Cells.ImportCustomObjects(
    items, props, true, 0, 0, items.Count, false, null, true);
```

## Exporting to a DataTable

`ExportDataTable(firstRow, firstColumn, totalRows, totalColumns, exportColumnName)`
infers one CLR type per column - it fails or coerces oddly when a column mixes types.
When a column holds mixed types, use `ExportDataTableAsString(...)` (same parameters,
every cell as `string`). An `ExportTableOptions` overload exposes `ExportColumnName`,
`ExportAsString`, `PlotVisibleRows`, and `SkipErrorValue`.

```csharp
Workbook wb = new Workbook("input.xlsx");
Cells cells = wb.Worksheets[0].Cells;
DataTable typed = cells.ExportDataTable(0, 0, 10, 3, true);         // strongly typed columns
DataTable text  = cells.ExportDataTableAsString(0, 0, 10, 3, true); // mixed columns -> string
Console.WriteLine(typed.Rows.Count + " / " + text.Rows.Count);
```

## Find and replace

`Cells.Find(object value, Cell previousCell, FindOptions options)` returns the matching
`Cell` or `null`; pass the previous hit back in to continue. Set `FindOptions.LookInType`
(`Values`, `Formulas`, `OriginalValues`, `FormattedValues`, ...) and
`FindOptions.LookAtType` (`Contains`, `StartWith`, `EndWith`, `EntireContent`); other
options include `CaseSensitive` and `RegexKey`. For blanket substitution,
`Worksheet.Replace(oldString, newString)` returns the count of changed cells.

```csharp
Workbook wb = new Workbook("input.xlsx");
Cells cells = wb.Worksheets[0].Cells;
FindOptions opts = new FindOptions();
opts.LookInType = LookInType.Values;
opts.LookAtType = LookAtType.EntireContent;  // whole-cell match; Contains for substring
int hits = 0;
Cell found = cells.Find("Oranges", null, opts);
while (found != null)
{
    hits++;
    found = cells.Find("Oranges", found, opts);   // pass previous hit to continue
}
Console.WriteLine(hits + " cells match");
int changed = wb.Worksheets[0].Replace("OldName", "NewName");
Console.WriteLine(changed + " cells replaced");
```

## Hyperlinks

`Worksheet.Hyperlinks.Add(cellName, totalRows, totalColumns, address)` links a cell or
range to a URL, another cell (`"Sheet1!B10"`), or an external file. If the cell is empty,
the address is written as the cell's text; if it already has a value, that value shows as
plain text and you must style it to look like a link (see styles-and-formatting.md).

```csharp
Workbook wb = new Workbook();
Worksheet sheet = wb.Worksheets[0];
sheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com"); // external URL
sheet.Hyperlinks.Add("A2", 1, 1, "Sheet1!B10");             // in-workbook target
sheet.Hyperlinks.Add("A3", 1, 1, "report.xlsx");            // external file
```

## Related

- styles-and-formatting.md - number/date display formats, aligning imported values, link styling.
- performance-and-large-files.md - streaming huge sheets, MemorySetting, LightCells import/export.
- formulas-and-calculation.md - reading computed results, CalculateFormula.
- worksheets-rows-columns.md - inserting/deleting rows and columns, ranges, validation.
