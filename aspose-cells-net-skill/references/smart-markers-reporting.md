# Smart Markers (Template Reporting)

Read when filling an Excel template whose cells contain `&=` markers from data (DataTable, objects, JSON), or when markers stay literal / data does not populate.

Smart markers are `&=`-prefixed placeholders in a designer workbook. `WorkbookDesigner` replaces them with data. A marker splits into a data-source segment and a field segment: `&=Source.Field`. Each row-marker inserts rows below it and pushes lower content down, so totals placed on the row immediately under a marker calculate over the inserted data.

## Core workflow

Put markers in a template (built in code or opened from disk), bind a source, process, then save the designer's OWN workbook:

```csharp
Workbook book = new Workbook();
Cells cells = book.Worksheets[0].Cells;
cells["A1"].PutValue("&=Orders.Customer");   // marker = &=Source.Field
cells["B1"].PutValue("&=Orders.Amount");

DataTable orders = new DataTable("Orders");  // TableName MUST equal marker source "Orders"
orders.Columns.Add("Customer");
orders.Columns.Add("Amount", typeof(double));
orders.Rows.Add("Acme", 1200.0);

WorkbookDesigner designer = new WorkbookDesigner(book);
designer.SetDataSource(orders);
designer.Process();
designer.Workbook.Save("report.xlsx");
```

Key facts:
- Constructors: `new WorkbookDesigner()` then assign `designer.Workbook = book`, or `new WorkbookDesigner(book)`. Both are valid.
- `SetDataSource(DataTable)` takes no name argument, so the marker's source segment is matched against `DataTable.TableName`. Name the table to exactly the marker source.
- `Process()` clears unrecognized markers. `Process(bool isPreserved)` with `true` leaves unrecognized markers in place. `Process(int sheetIndex, bool isPreserved)` processes a single sheet (use it per sheet when data overflows to a second sheet that repeats the same markers).
- Always save `designer.Workbook` - the instance you processed - not a separate `Workbook`.

## Pitfalls

### Markers stay literal after Process (cells still show `&=Orders.Customer`)

Symptom: output still contains the raw `&=...` text; no data lands.
Cause (one of): the source name does not equal the marker's source segment; a typo (`&` instead of `&=`, or a space after `&=`); `SetDataSource` was never called for that source name; `Process()` was never called; or a different `Workbook` instance was saved than the one processed.
Fix: set `DataTable.TableName` (or the `SetDataSource(name, ...)` name) to exactly the marker source; call `Process()`; save `designer.Workbook`. A `new DataTable()` with no name never matches - name it.

### Only the first row populates, or inserted rows overwrite content below

Symptom: one data row appears; rows beneath the marker get clobbered, or unwanted blank rows pile up.
Cause: fighting the default. By default a row-marker inserts a fresh row per record and pushes lower content down; `noadd` suppresses insertion so data overwrites whatever already sits below.
Fix: leave insertion on (do NOT add `noadd`) so the sheet grows to fit the data. Reserve `noadd`+`skip:n` for deliberate alternating-row layouts, and put `noadd` on the TOP marker because templates process bottom-to-top. Add `copystyle` to carry the marker cell's style down every inserted row.

## Marker syntax

| Marker | Meaning |
| - | - |
| `&=Source.Field` | Field of a bound table/collection; expands one row per record. |
| `&=[Data Source].[Field Name]` | Bracket form when a name contains spaces. |
| `&=$VarName` | Scalar variable - fills exactly one cell. |
| `&=$VarArray` | Variable array - fills several cells. |
| `&=&=B{r}*C{r}` | Repeating dynamic formula; `{r}` = current row, `{-1}`/`{+2}` = row offsets. |
| `&=Source.Obj.Sub` | Nested object property (keep the nesting simple). |
| `&=Source.Photo(Picture:FitToCell)` | Image marker (see Images). |

Only one marker per cell. Unused markers are removed on `Process()`.

## Marker parameters

Append in parentheses, comma-separated, no spaces: `&=Orders.Amount(copystyle,skip:1)`.

| Parameter | Meaning |
| - | - |
| `noadd` | Do not insert rows to fit data (overwrites existing rows). |
| `skip:n` | Skip n rows after each data row. |
| `ascending:n` / `descending:n` | Sort by this column; n = its order among sort keys. |
| `horizontal` | Write left-to-right instead of top-to-bottom. |
| `numeric` | Convert text to number where possible. |
| `shift` | Insert cells down/right to fit, like Excel Insert with shift. |
| `copystyle` | Copy the marker cell's style to all cells in the column. |
| `formula` | Treat the field value as a real formula. |
| `group:normal` / `group:merge` / `group:repeat` | Group rows by this field (see Grouping). |
| `subtotalN:Ref` | Summary of this field per group `Ref` (see Grouping). |
| `picture` | Insert the field value as a picture. |

## Grouping and subtotals

Group rows by a field and emit summary rows between groups:
- `group:normal` prints the group value once per group; `group:merge` does the same and merges the group cells; `group:repeat` repeats it on every row.
- `subtotalN:Table.GroupCol` runs Excel's SUBTOTAL function N per group (1=AVERAGE, 2=COUNT, 3=COUNTA, 4=MAX, 5=MIN, 9=SUM, and so on). Group by several columns with `&`: `subtotal9:Table1.ColA&Table1.ColB`.

Example markers on the data row, over an `Order Details` source:
`&=[Order Details].OrderID(group:merge,skip:1)` and `&=[Order Details].Quantity(subtotal9:Order Details.OrderID)`.
A standalone summary cell one row under the data: `&=subtotal9:Order Details.OrderID`.
Add `Label` / `LabelPosition` inside the subtotal marker for a custom total label.

## Custom objects (POCO collection)

Bind a `List<T>` (or array) with the named `SetDataSource(string, object)` overload; each element becomes a row and the name you pass is the marker source.

```csharp
public class Employee
{
    public string Name { get; set; }
    public double Salary { get; set; }
}
```

```csharp
List<Employee> staff = new List<Employee>
{
    new Employee { Name = "Ann", Salary = 5000 },
    new Employee { Name = "Bob", Salary = 6200 }
};
WorkbookDesigner designer = new WorkbookDesigner(new Workbook("template.xlsx"));
designer.SetDataSource("Employee", staff);   // markers: &=Employee.Name, &=Employee.Salary
designer.Process();
designer.Workbook.Save("report.xlsx");
```

Anonymous types and simple nested properties (`&=Employee.Address.City`) bind the same way. For a fully custom row iterator (lazy or proprietary data), implement `ICellsDataTable` and pass it via `SetDataSource(string, ICellsDataTable)`.

## JSON data source

Pass raw JSON text with `SetJsonDataSource(string name, string json)`. If the JSON root is an object, `name` may be `null` and its array properties become the marker sources; if the root is a bare array, give it a name.

```csharp
string json = File.ReadAllText("table.json");  // { "Items": [ { "ItemName": "A123", "Qty": 55 } ] }
WorkbookDesigner designer = new WorkbookDesigner(new Workbook("template.xlsx"));
designer.SetJsonDataSource(null, json);         // markers: &=Items.ItemName, &=Items.Qty
designer.Process();
designer.Workbook.CalculateFormula();
designer.Workbook.Save("report.xlsx");
```

## Scalar variables and dynamic formulas

- Scalar variable: `designer.SetDataSource("Year", 2026);` fills `&=$Year` (one cell). A variable array fills consecutive cells.
- Dynamic formula: put a repeating formula marker so it references rows that get inserted, e.g. `cells["C1"].PutValue("&=&=A{r}*B{r}");` yields `=A1*B1`, `=A2*B2`, ... Append parameters after a `~` separator: `&=&=B{-1}/C{-1}~(skip:1)`.
- Formula parameter: `&=Src.Field(formula)` writes each field's string value as a live formula instead of literal text.

Dynamic and `formula`-parameter markers produce formulas, not computed values. Set `designer.CalculateFormula = true` before `Process()`, or call `designer.Workbook.CalculateFormula()` afterward, before reading the computed cell values.

## Images

Image markers render a field that holds image data as a picture in the cell. Syntax: `&=Source.Photo(Picture:FitToCell)`. Options: `Picture:FitToCell` (fit to the row height and column width), `Picture:ScaleN` (scale height and width to N percent), `Picture:Width:Nin & Height:Nin` (size in inches). Combines with grouping, e.g. `&=Person.Photo(Picture:FitToCell)` alongside `&=Person.Name(group:normal,skip:1)`.

## After Process

Computed values need calculation before you read them: `designer.Workbook.CalculateFormula()` (see formulas-and-calculation.md). Then save or convert the workbook to any format (see conversion-and-rendering.md). For thousands of records or many templates in one run, see performance-and-large-files.md.

## Related
- cell-values-and-data.md - building DataTables and general cell IO.
- formulas-and-calculation.md - CalculateFormula after Process for computed cells.
- styles-and-formatting.md - formatting the populated report.
- conversion-and-rendering.md - convert the finished report to PDF or other formats.
- setup-and-licensing.md - apply the license before processing (watermark otherwise).
- performance-and-large-files.md - memory settings for very large generated reports.
