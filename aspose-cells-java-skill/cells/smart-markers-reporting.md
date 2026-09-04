---
name: smart-markers-reporting
description: Fill an Excel template whose cells contain `&=` smart markers from an `ICellsDataTable` or a JSON string using Aspose.Cells for Java WorkbookDesigner. Use when the user provides a designer XLSX with `&=Source.Field` placeholders and asks to bind data, or when the result keeps the raw `&=` text and no data lands.
applies_to: Java
keywords: [smart markers, &=, WorkbookDesigner, template, data binding, JSON, ICellsDataTable]
---

# Smart markers (template reporting)

Smart markers are `&=`-prefixed placeholders in a designer workbook
(built in code or loaded from disk). `WorkbookDesigner` replaces them
with data. Each row-marker inserts rows below it and pushes lower
content down, so a totals row immediately under a marker calculates
over the inserted data.

Aspose.Cells for Java exposes `com.aspose.cells.WorkbookDesigner`.
There is no `java.util.DataTable` — the cross-language equivalent of
.NET's `DataTable` is `com.aspose.cells.ICellsDataTable`. Implement
it (or wrap a `Collection` via `CellsDataTableFactory`) and pass the
result as the source.

## Imports

```java
import com.aspose.cells.*;
```

## Core workflow

```java
Workbook book = new Workbook();
Worksheet ws = book.getWorksheets().get(0);
ws.getCells().get("A1").setValue("&=Orders.Customer");   // marker = &=Source.Field
ws.getCells().get("B1").setValue("&=Orders.Amount");

// Inline ICellsDataTable over a list of {Customer -> Object} maps.
final String[] cols = {"Customer", "Amount"};
final java.util.List<java.util.Map<String, Object>> rows = new java.util.ArrayList<>();
java.util.Map<String, Object> r = new java.util.LinkedHashMap<>();
r.put("Customer", "Acme");
r.put("Amount", 1200.0);
rows.add(r);

ICellsDataTable data = new ICellsDataTable() {
    int idx = -1;
    public String[] getColumns()      { return cols; }
    public int    getCount()          { return rows.size(); }
    public void   beforeFirst()       { idx = -1; }
    public Object get(int column)     { return rows.get(idx).get(cols[column]); }
    public Object get(String name)    { return rows.get(idx).get(name); }
    public boolean next()             { idx++; return idx < rows.size(); }
};

WorkbookDesigner designer = new WorkbookDesigner(book);
designer.setDataSource("Orders", data);                // "Orders" must match marker source
designer.process();
designer.getWorkbook().save("report.xlsx");
```

Key points:

- The first argument to `setDataSource(String, ...)` **must equal**
  the marker's source segment — `&=Orders.Customer` needs a source
  named `Orders`.
- `process()` clears unrecognized markers. `process(boolean
  isPreserved)` with `true` leaves them in place.
- Per-sheet processing: `process(int sheetIndex, boolean isPreserved)`
  (use when data overflows to a second sheet that repeats markers).
- Always save `designer.getWorkbook()` — the instance you processed —
  not a separate `Workbook`.

## Pitfalls

### Markers stay literal after `process()` (cells still show `&=Orders.Customer`)

Symptom: the saved file still contains the raw `&=...` text; no data
lands. Cause (one of):

- The source name passed to `setDataSource(...)` does not match the
  marker's source segment (`Orders` vs `orders`).
- A typo: `&` instead of `&=`, or a space after `&=`.
- `setDataSource(...)` was never called for the source name.
- `process()` was never called.
- A different `Workbook` instance was saved than the one processed.

Fix: pass the exact source string (`"Orders"` for
`&=Orders.Customer`); call `process()`; save `designer.getWorkbook()`.

```java
Workbook book = new Workbook();
Worksheet ws = book.getWorksheets().get(0);
ws.getCells().get("A1").setValue("&=Orders.Customer");
ws.getCells().get("B1").setValue("&=Orders.Amount");

final String[] cols = {"Customer", "Amount"};
final java.util.List<java.util.Map<String, Object>> rows = new java.util.ArrayList<>();
java.util.Map<String, Object> r = new java.util.LinkedHashMap<>();
r.put("Customer", "Acme");
r.put("Amount", 1200.0);
rows.add(r);

ICellsDataTable data = new ICellsDataTable() {
    int idx = -1;
    public String[] getColumns()      { return cols; }
    public int    getCount()          { return rows.size(); }
    public void   beforeFirst()       { idx = -1; }
    public Object get(int column)     { return rows.get(idx).get(cols[column]); }
    public Object get(String name)    { return rows.get(idx).get(name); }
    public boolean next()             { idx++; return idx < rows.size(); }
};

WorkbookDesigner designer = new WorkbookDesigner(book);
designer.setDataSource("Orders", data);
designer.process();
designer.getWorkbook().save("report.xlsx");
```

### Only the first row populates, or inserted rows clobber content below

Symptom: one data row appears, or rows beneath the marker get
clobbered, or unwanted blank rows pile up.

Cause: fighting the default. By default a row-marker inserts a fresh
row per record and pushes lower content down; `noadd` suppresses
insertion so data overwrites whatever already sits below.

Fix: leave insertion on (do NOT add `noadd`) so the sheet grows to
fit the data. Reserve `noadd` + `skip:n` for deliberate
alternating-row layouts, and put `noadd` on the TOP marker because
templates process bottom-to-top. Add `copystyle` to carry the marker
cell's style down every inserted row.

```java
// Markers on row 1; row 2 below them is the dynamic-formula totals row.
// The template processes bottom-to-top, so totals placed under the data
// row reference data inserted above via {r}:
Workbook book = new Workbook();
Worksheet ws = book.getWorksheets().get(0);
ws.getCells().get("A1").setValue("&=Orders.Customer");
ws.getCells().get("B1").setValue("&=Orders.Amount(copystyle)");
ws.getCells().get("C2").setValue("&=&=SUM(B{r})");   // dynamic formula
```

### Formula cells stay blank after `process()`

Symptom: total or summary rows beneath inserted data show 0 instead
of the calculated value.

Cause: marker processing inserts data but does not recompute formulas.
Fix: turn on `setCalculateFormula(true)` on the designer or call
`workbook.calculateFormula()` once before save.

```java
Workbook book = new Workbook();
Worksheet ws = book.getWorksheets().get(0);
ws.getCells().get("A1").setValue("&=Orders.Customer");
ws.getCells().get("B1").setValue("&=Orders.Amount");

final String[] cols = {"Customer", "Amount"};
final java.util.List<java.util.Map<String, Object>> rows = new java.util.ArrayList<>();
java.util.Map<String, Object> r = new java.util.LinkedHashMap<>();
r.put("Customer", "Acme");
r.put("Amount", 1200.0);
rows.add(r);

ICellsDataTable data = new ICellsDataTable() {
    int idx = -1;
    public String[] getColumns()      { return cols; }
    public int    getCount()          { return rows.size(); }
    public void   beforeFirst()       { idx = -1; }
    public Object get(int column)     { return rows.get(idx).get(cols[column]); }
    public Object get(String name)    { return rows.get(idx).get(name); }
    public boolean next()             { idx++; return idx < rows.size(); }
};

WorkbookDesigner designer = new WorkbookDesigner(book);
designer.setDataSource("Orders", data);
designer.setCalculateFormula(true);          // wb.calculateFormula() runs during process()
designer.process();
designer.getWorkbook().save("report.xlsx");
```

Alternatively, recompute explicitly after `process()` instead of
`setCalculateFormula(true)`:

```java
Workbook book = new Workbook();
Worksheet ws = book.getWorksheets().get(0);
ws.getCells().get("A1").setValue("&=Orders.Customer");
ws.getCells().get("B1").setValue("&=Orders.Amount");

final String[] cols = {"Customer", "Amount"};
final java.util.List<java.util.Map<String, Object>> rows = new java.util.ArrayList<>();
java.util.Map<String, Object> r = new java.util.LinkedHashMap<>();
r.put("Customer", "Acme");
r.put("Amount", 1200.0);
rows.add(r);

ICellsDataTable data = new ICellsDataTable() {
    int idx = -1;
    public String[] getColumns()      { return cols; }
    public int    getCount()          { return rows.size(); }
    public void   beforeFirst()       { idx = -1; }
    public Object get(int column)     { return rows.get(idx).get(cols[column]); }
    public Object get(String name)    { return rows.get(idx).get(name); }
    public boolean next()             { idx++; return idx < rows.size(); }
};

WorkbookDesigner designer = new WorkbookDesigner(book);
designer.setDataSource("Orders", data);
designer.process();
designer.getWorkbook().calculateFormula();  // recompute formulas before save
designer.getWorkbook().save("report.xlsx");
```

## Marker syntax

| Marker                              | Meaning                                                       |
| ----------------------------------- | ------------------------------------------------------------- |
| `&=Source.Field`                    | Field of a bound table/collection; expands one row per record |
| `&=[Data Source].[Field Name]`      | Bracket form when a name contains spaces                      |
| `&=$VarName`                        | Scalar variable — fills exactly one cell                      |
| `&=$VarArray`                       | Variable array — fills several cells                          |
| `&=&=B{r}*C{r}`                     | Repeating dynamic formula; `{r}` = current row                |
| `&=Source.Obj.Sub`                  | Nested object property                                        |
| `&=Source.Photo(Picture:FitToCell)` | Image marker                                                  |

Only one marker per cell. Unused markers are removed on `process()`.

## Marker parameters

Append in parentheses, comma-separated, no spaces:
`&=Orders.Amount(copystyle,skip:1)`.

| Parameter                         | Meaning                                                    |
| --------------------------------- | ---------------------------------------------------------- |
| `noadd`                           | Do not insert rows to fit data (overwrites existing rows)  |
| `skip:n`                          | Skip n rows after each data row                            |
| `ascending:n` / `descending:n`    | Sort by this column; n = its order among sort keys         |
| `horizontal`                      | Write left-to-right instead of top-to-bottom               |
| `numeric`                         | Convert text to number where possible                      |
| `shift`                           | Insert cells down/right to fit, like Excel Insert with shift |
| `copystyle`                       | Copy the marker cell's style to all cells in the column    |
| `formula`                         | Treat the field value as a real formula                    |
| `group:normal` / `merge` / `repeat` | Group rows by this field                                 |
| `subtotalN:Ref`                   | Summary of this field per group (N: 1=AVERAGE, 9=SUM, ...) |
| `picture`                         | Insert the field value as a picture                        |

## Custom row iterators (lazy data, joined sources)

Implement `ICellsDataTable` for fully custom row iteration — lazy
loading, joined data, etc. The interface requires six methods:

- `getColumns()` — return the marker field names in stable order.
- `getCount()` — return the row count, or `-1` for unknown/lazy data.
- `beforeFirst()` — rewind the cursor to before the first row.
- `next()` — advance; return `true` if a row is now active, `false`
  to signal end of data.
- `get(int columnIndex)` and `get(String columnName)` — return the
  column value from the current row.

For a fully managed POJO `Collection`, wrap it via the internal
`CellsDataTableFactory.getInstance(Collection)` (the
`CellsDataTableFactory` constructor is internal) or, easier, use
the anonymous-class `ICellsDataTable` pattern above.

## JSON data source

Pass raw JSON text with `setJsonDataSource(String, String)`. If the
JSON root is an object, `name` may be `null` and its array properties
become marker sources; if the root is a bare array, give it a name.

```java
WorkbookDesigner designer = new WorkbookDesigner(new Workbook("template.xlsx"));
designer.setJsonDataSource(null, "{\"Items\":[{\"ItemName\":\"A123\",\"Qty\":55}]}");
// markers: &=Items.ItemName / &=Items.Qty
designer.process();
designer.getWorkbook().calculateFormula();
designer.getWorkbook().save("report.xlsx");
```

## Dynamic formulas and variable markers

- Scalar variable: `designer.setDataSource("Year", 2026);` fills
  `&=$Year` (one cell). Variable arrays fill consecutive cells.
- Dynamic formula: `&=&=A{r}*B{r}` yields `=A1*B1`, `=A2*B2`, ...
  across the inserted rows. Append parameters after a `~` separator:
  `&=&=B{-1}/C{-1}~(skip:1)`.
- `formula`-parameter marker: `&=Src.Field(formula)` writes each
  field's value as a live formula instead of literal text.

Dynamic and `formula`-parameter markers produce formulas, not
computed values. Set `setCalculateFormula(true)` on the designer, or
call `workbook.calculateFormula()` afterward, before reading computed
cell values.

## After `process()`

Computed values need calculation before you read or convert them;
`designer.getWorkbook().calculateFormula()` does it. Save to any
output format (`designer.getWorkbook().save(...)` plus `SaveOptions`)
or convert via `PdfSaveOptions` (see `rendering/convert-to-pdf.md`).
For many templates in one run, see
`_shared/large-files-memory.md`.

## Related

- `cells/read-write-values.md` — basic cell IO and writing Collections.
- `_shared/large-files-memory.md` — memory settings for very large generated reports.
- `_shared/license.md` — apply a license before processing (watermark otherwise).
- `rendering/convert-to-pdf.md` — convert the finished report to PDF.
