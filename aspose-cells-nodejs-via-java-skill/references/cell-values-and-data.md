# Cell values and data

## Typed reads and writes

Use `get(row, column)` with 0-based indices or `get("A1")`. `putValue()` handles common primitive values without constructing a cell object yourself.

```javascript
const cellsApi = require("aspose.cells");

const workbook = new cellsApi.Workbook();
const cells = workbook.getWorksheets().get(0).getCells();
cells.get("A1").putValue("Quantity");
cells.get("B1").putValue(12);
cells.get("C1").putValue(19.95);
cells.get("D1").putValue(true);

const quantity = cells.get("B1").getValue();
console.log(quantity, cells.get("B1").getType());
```

Use `getStringValue()`, `getDoubleValue()`, or `getBoolValue()` only when the cell's actual type is compatible. Check `getType()` against `CellValueType` when reading untrusted mixed data. Do not silently stringify formulas or numbers before applying formatting.

## Dates

A JavaScript `Date` is not guaranteed to bridge to Aspose's Java `DateTime` value. Prefer a deliberate string representation and number format when that satisfies the requirement:

```javascript
const cellsApi = require("aspose.cells");

const workbook = new cellsApi.Workbook();
const cell = workbook.getWorksheets().get(0).getCells().get("A1");
cell.putValue("2026-11-23");
const style = cell.getStyle();
style.setCustom("yyyy-mm-dd");
cell.setStyle(style);
```

If arithmetic requires a true Excel date/time value, construct the `DateTime` representation explicitly documented by the installed Node via Java release and test the round trip. Never rely on an implicit `new Date()` conversion.

## Write rows without Java-array APIs

The Java-style array and `ResultSet` import APIs are limitations of this binding. In particular, do not pass a JavaScript array to `importArray`, `importObjectArray`, or related Java-array overloads. A small data import is clearer and more reliable as typed writes:

```javascript
const cellsApi = require("aspose.cells");

const rows = [
  ["North", 10, 125.5],
  ["South", 8, 99.0]
];

const workbook = new cellsApi.Workbook();
const cells = workbook.getWorksheets().get(0).getCells();
for (let r = 0; r < rows.length; r += 1) {
  for (let c = 0; c < rows[r].length; c += 1) {
    cells.get(r, c).putValue(rows[r][c]);
  }
}
```

For a large import, reuse styles, avoid unnecessary per-cell object creation, and use the memory guidance in [performance](performance-and-large-files.md). Only use a bulk API after proving that its exact parameter type is supported by the installed bridge; a Java-compatible collection is not the same thing as a JavaScript array.

## Iterate existing cells

Avoid scanning a large rectangle with `cells.get(r, c)`: accessing missing cells can instantiate them and grow the workbook. Iterate materialized rows and their cells instead.

```javascript
const cellsApi = require("aspose.cells");

const workbook = new cellsApi.Workbook("input.xlsx");
const rows = workbook.getWorksheets().get(0).getCells().getRows();
for (let r = 0; r < rows.getCount(); r += 1) {
  const row = rows.get(r);
  for (let i = 0; i < row.getCount(); i += 1) {
    const cell = row.getCellByIndex(i);
    console.log(cell.getName(), cell.getValue());
  }
}
```

Use explicit per-cell reads for exports whose destination must be a plain JavaScript array. Keep formulas and displayed values distinct: `getValue()` returns the cached value, which can be stale until recalculation.
