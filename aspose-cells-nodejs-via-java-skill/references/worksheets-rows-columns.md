# Worksheets, rows, and columns

All sheet, row, and column indices are 0-based.

## Add, copy, and remove sheets

Collection `add()` methods commonly return an index; retrieve the object from the collection rather than treating the return value as a worksheet.

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook();
const sheets = workbook.getWorksheets();
sheets.get(0).setName("Source");

const addedIndex = sheets.add();
sheets.get(addedIndex).setName("Output");
const copiedIndex = sheets.addCopy("Source");
sheets.get(copiedIndex).setName("Source Copy");
sheets.removeAt("Source Copy");
```

Check sheet-name validity and duplicates before saving. Keep at least one visible worksheet in a workbook.

## Insert and delete structural elements

Insert/delete APIs update dependent references and formulas themselves. Do not manually rewrite formulas a second time.

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook("input.xlsx");
const sheet = workbook.getWorksheets().get(0);
sheet.getCells().insertRows(1, 2);
sheet.getCells().deleteRows(5, 1);
sheet.getCells().insertColumns(2, 1);
sheet.getCells().deleteColumns(4, 1);
sheet.autoFitColumn(0);
```

When deleting several unrelated rows by index, iterate bottom-up so earlier deletions do not change the indices still to process. Recalculate formulas after structural edits before reading cached results.

## Ranges and tables

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook();
const sheet = workbook.getWorksheets().get(0);
const rows = [
  ["Region", "Amount"],
  ["North", 125],
  ["South", 99]
];
for (let r = 0; r < rows.length; r += 1) {
  for (let c = 0; c < rows[r].length; c += 1) {
    sheet.getCells().get(r, c).putValue(rows[r][c]);
  }
}

const tableIndex = sheet.getListObjects().add(0, 0, 2, 1, true);
sheet.getListObjects().get(tableIndex).setDisplayName("SalesTable");
sheet.autoFitColumn(0);
sheet.autoFitColumn(1);
```

A table header row must contain appropriate header values. Keep table display names unique and free of spaces.

## Data validation

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook();
const sheet = workbook.getWorksheets().get(0);
const area = new cells.CellArea();
area.setStartRow(0);
area.setEndRow(99);
area.setStartColumn(1);
area.setEndColumn(1);

const index = sheet.getValidations().add(area);
const validation = sheet.getValidations().get(index);
validation.setType(cells.ValidationType.LIST);
validation.setFormula1('"Open,Closed"');
```

Use worksheet protection only after setting the desired editable cells and protection options. Protection is not encryption; use file encryption when confidentiality is required.
