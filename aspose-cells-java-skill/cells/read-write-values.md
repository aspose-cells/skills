---
name: read-write-values
description: Read and write cell values in an Aspose.Cells for Java workbook — strings, numbers, dates, formulas, and rich-text, by cell reference (A1) or index. Use whenever the user wants to set or read cell content.
applies_to: Java
keywords: [cell, value, read, write, getValue, setValue, a1, range]
---

# Read & write cell values

## Imports

```java
import com.aspose.cells.*;
```

## Quick example

```java
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().get(0);
Cells cells = ws.getCells();

// By reference
cells.get("A1").setValue("Customer");
cells.get("B1").setValue(1234.56);

// Read back
String header = cells.get("A1").getStringValue();
double total  = cells.get("B1").getDoubleValue();
```

## All supported `putValue` overloads on `Cell`

| Overload signature               | Use                              |
| -------------------------------- | -------------------------------- |
| `putValue(Object value)`         | dispatcher — rare                |
| `putValue(String value)`         | text cell                        |
| `putValue(double value)`         | numeric cell                     |
| `putValue(int value)`            | integer numeric                  |
| `putValue(boolean value)`        | boolean cell                     |
| `putValue(com.aspose.cells.DateTime)` | date (paired with date style) |

```java
Cell cell = ws.getCells().get("A1");
cell.putValue("text");
cell.putValue(123.45);
cell.putValue(42);
cell.putValue(true);
cell.putValue(new com.aspose.cells.DateTime(2024, 1, 15));
```

Reading:

| Method                              | Returns            |
| ----------------------------------- | ------------------ |
| `getValue()`                        | `Object`           |
| `getStringValue()`                  | `String`           |
| `getDoubleValue()`                  | `double`           |
| `getIntValue()`                     | `int`              |
| `getBoolValue()`                    | `boolean`          |
| `getDateTimeValue()`                | `java.util.Date`   |

## Insert data with PutValue

```java
Worksheet ws = wb.getWorksheets().get(0);
Cells cells = ws.getCells();
cells.get("A2").putValue("Total");
cells.get("B2").putValue(50.0);
```

`putValue` and `setValue` are equivalent for strings/numbers/booleans;
prefer `setValue` to keep one mental model.

## Bulk insert — arrays of strings

```java
Worksheet ws = wb.getWorksheets().get(0);
Cells cells = ws.getCells();
String[] names = {"Alice", "Bob", "Carol"};
cells.importArray(names, 0, 0, true);    // firstRow, firstCol, isVertical
```

## Iterate cells (avoid `for-each` on `Cells`)

```java
Worksheet ws = wb.getWorksheets().get(0);
Cells cells = ws.getCells();
for (int r = cells.getMinRow(); r <= cells.getMaxRow(); r++) {
    for (int c = cells.getMinColumn(); c <= cells.getMaxColumn(); c++) {
        Cell cell = cells.checkCell(r, c);
        if (cell == null) continue;
        // process
    }
}
```

`checkCell` returns `null` for empty cells — much cheaper than
`getCell`, which creates empty `Cell` instances.

## Pitfalls

- **Use JavaBean accessors only** — `cell.getValue()` /
  `cell.putValue(...)`. Never `cell.Value` (C#).
- **`putValue("123")` stores as a string**, not a number. To force
  numeric, pass `double`/`int` directly: `cell.putValue(123)`.
- **`putValue(date)` stores as a date serial only if the cell style
  has a date format.** Otherwise Excel displays it as a number.
- **Strings with a leading apostrophe** (e.g. `'00123`) —
  `putValue("00123")` strips the quote. To preserve it, set
  `QuotePrefix` on the cell's `Style` —
  `cell.getStyle().setQuotePrefix(true); cell.setStyle(s);`
  (mutate a copy and assign back; see `cell-formatting-styles.md`).
- **Formula strings** must start with `=` — see
  `formulas/set-get-formula.md`.

### `getValue()` returns an unpredictable runtime type

Symptom: `(double) cell.getValue()` throws `ClassCastException`; a
date cell comes back as a `Double`; an integer sometimes boxes as
`Integer`, sometimes as `Double`.

Cause: `Cell.getValue()` is `Object`. Its runtime type is one of
`null`, `Boolean`, `com.aspose.cells.DateTime`, `Double`, `Integer`,
or `String` — and Aspose does not guarantee a stable type across
calls or versions. Date cells are stored as OADate serials, so a
date cell may come back as `DateTime` OR as a `Double`.

Fix: branch on `Cell.getType()` and read through the typed
accessors. `CellValueType` is an int-constants class with
`IS_NUMERIC`, `IS_DATE_TIME`, `IS_STRING`, `IS_BOOL`, `IS_ERROR`,
`IS_NULL`, `IS_UNKNOWN`.

```java
Cell cell = ws.getCells().get("A1");
switch (cell.getType()) {
    case CellValueType.IS_NUMERIC:  System.out.println(cell.getDoubleValue());  break;
    case CellValueType.IS_DATE_TIME: System.out.println(cell.getDateTimeValue()); break;
    case CellValueType.IS_BOOL:     System.out.println(cell.getBoolValue());     break;
    case CellValueType.IS_STRING:   System.out.println(cell.getStringValue());   break;
    case CellValueType.IS_ERROR:    System.out.println(cell.getStringValue());   break; // "#DIV/0!"
    case CellValueType.IS_NULL:     break;                                          // empty
}
```

`getDoubleValue()` / `getIntValue()` throw on an empty
(`IS_NULL`) cell — guard first.

### Read string differs from what Excel shows

Symptom: a cell displays `0.01` in Excel but your read returns
`0.012345`, or vice versa; you cannot tell whether to use the
underlying value or the formatted display.

Cause: three different strings exist for one cell.
`getStringValue()` applies the cell's own number format (identical
to Excel "copy as text"). `getDisplayStringValue()` applies the
cell's display style — exactly what Excel renders, including
conditional-formatting overrides. `getStringValueWithoutFormat()`
returns the raw underlying value.

Fix: pick the accessor that matches intent. For explicit control
use the `getStringValue(CellValueFormatStrategy)` overload —
`CELL_STYLE` = cell's own number format, `DISPLAY_STYLE` = as
displayed in Excel, `DISPLAY_STRING` = same as `getDisplayStringValue`,
`NONE` = raw.

```java
Cell cell = ws.getCells().get("A1");                       // 0.012345, format "0.00"
System.out.println(cell.getStringValue());                 // "0.01"     - cell's format
System.out.println(cell.getDisplayStringValue());          // "0.01"     - what Excel shows
System.out.println(cell.getStringValue(CellValueFormatStrategy.NONE)); // "0.012345" - raw
```

### Text numbers stay text and break SUM

Symptom: imported/typed numbers are left-aligned, `SUM` ignores
them, sorting is lexical.

Cause: a string like `"00123"` written as `putValue` stays text —
neither `putValue(String)` nor `setValue(...)` auto-parses.

Fix: parse on write — pass a numeric type (`putValue(123.0)`), or
fix a whole sheet after load with `Cells.convertStringToNumericValue()`.

```java
Cells cells = ws.getCells();
cells.get("A1").putValue("00123");           // stays text → left-aligned, excluded from SUM
cells.get("A2").putValue(123.0);             // numeric
cells.convertStringToNumericValue();          // whole sheet: parse every text-number in place
```

### Bulk import — `ImportTableOptions` instead of `importArray` for tabular data

For a `DataTable` / `ICellsDataTable` / `Object[][]` with a header
row and typed columns, use `Cells.importData(...)` with an
`ImportTableOptions`. Configure behaviour via the options
instance (header row, parse numeric strings, date format, etc.):

```java
// Anonymous ICellsDataTable — same pattern as cells/smart-markers-reporting.md
final String[] cols = {"Name", "Price"};
final java.util.List<java.util.Map<String, Object>> rows = new java.util.ArrayList<>();
java.util.Map<String, Object> r = new java.util.LinkedHashMap<>();
r.put("Name", "Pen"); r.put("Price", 1.5);
rows.add(r);

ICellsDataTable data = new ICellsDataTable() {
    int idx = -1;
    public String[] getColumns()  { return cols; }
    public int getCount()          { return rows.size(); }
    public void beforeFirst()      { idx = -1; }
    public Object get(int column)  { return rows.get(idx).get(cols[column]); }
    public Object get(String name) { return rows.get(idx).get(name); }
    public boolean next()          { idx++; return idx < rows.size(); }
};

ImportTableOptions opts = new ImportTableOptions();
opts.setFieldNameShown(true);           // header row
opts.setConvertNumericData(true);       // parse numeric-looking strings
opts.setDateFormat("yyyy-MM-dd");       // dates format
ws.getCells().importData(data, 0, 0, opts);
```

`ImportTableOptions` writes values only — apply date/number display
formats afterward (see `cell-formatting-styles.md`). Other
load-tiny helpers: `Cells.importArray(...)` for 1D typed arrays,
`Cells.importTwoDimensionArray(...)` for a rectangular `Object[,]`.

## Related

- [formulas/set-get-formula.md](../formulas/set-get-formula.md)
- [cell-formatting-styles.md](cell-formatting-styles.md)
- [rich-text.md](rich-text.md)
