# Worksheets, Rows, Columns, Ranges, Tables, Protection

Read this when adding/copying/moving/deleting sheets, inserting/deleting/hiding rows or columns, autofitting, named ranges, Excel tables (ListObjects), data validation, freeze panes, or protecting workbooks/worksheets/cells.

All indices are 0-based: row 0 is Excel row 1, column 0 is column A. The workhorse is `Worksheet.Cells`.

## Pitfalls

### Off-by-one and self-shifting indices when inserting/deleting

Symptom: content lands one row/column off, or a delete-in-a-loop removes the wrong rows and skips some.

Cause: indices are 0-based (Excel row 1 = index 0). `InsertRows`/`DeleteRows` shift ALL existing content and auto-update formulas, chart sources, and references; deleting index `r` renumbers every row below it, so a forward loop targets the wrong rows after the first pass.

Fix: never manually re-adjust references after an insert/delete - Aspose does it for you. Delete a contiguous block in ONE call (`cells.DeleteRows(2, 5)`). For scattered rows, iterate descending so earlier deletes cannot shift indices you have not processed yet.

```csharp
Workbook wb = new Workbook("input.xlsx");
Cells cells = wb.Worksheets[0].Cells;
// Delete scattered rows 2, 5, 9 (0-based) top-safe: process DESCENDING.
int[] rowsToDelete = { 9, 5, 2 };
foreach (int r in rowsToDelete)
    cells.DeleteRows(r, 1);
wb.Save("output.xlsx");
```

`InsertRows(rowIndex, totalRows)` inserts above `rowIndex` and pushes content down; `InsertColumn(columnIndex)` pushes right. Overloads add `updateReference` (bool) or `InsertOptions`/`DeleteOptions`.

### AutoFit is slow or "does nothing"

Symptom: `AutoFitColumns`/`AutoFitRows` takes seconds on large sheets, or leaves merged/wrapped cells unsized.

Cause: autofit measures every cell in scope - it is O(cells) and expensive. By design it ignores merged cells and wrapped-text columns unless you pass options.

Fix: call it ONCE after all data is written, never per-row. Narrow the scope with `AutoFitColumns(firstColumn, lastColumn)`. For merged cells pass `AutoFitterOptions`.

```csharp
Workbook wb = new Workbook("input.xlsx");
Worksheet sheet = wb.Worksheets[0];
// ... write ALL data first, then autofit once ...
AutoFitterOptions opts = new AutoFitterOptions();
opts.AutoFitMergedCellsType = AutoFitMergedCellsType.EachLine;
opts.OnlyAuto = true;              // skip rows/cols the user sized manually
sheet.AutoFitColumns(opts);
sheet.AutoFitRows(opts);
wb.Save("output.xlsx");
```

`AutoFitMergedCellsType`: None, FirstLine, LastLine, EachLine. Other options: `IgnoreHidden`, `MaxRowHeight`, `AutoFitWrappedTextType`.

### FreezePanes argument order

Symptom: the wrong rows freeze, or nothing looks frozen.

Cause: misreading `FreezePanes(int row, int column, int freezedRows, int freezedColumns)`. The first pair is the top-left cell of the SCROLLABLE area (where the split sits); the second pair is HOW MANY rows/columns stay frozen.

Fix: to freeze only the top row, split at A2 (row 1, col 0) and freeze 1 row, 0 columns.

```csharp
Workbook wb = new Workbook("input.xlsx");
Worksheet sheet = wb.Worksheets[0];
sheet.FreezePanes(1, 0, 1, 0);     // freeze the top row
// Freeze first two rows AND first column: FreezePanes(2, 1, 2, 1)
wb.Save("output.xlsx");
```

`FreezePanes(string cellName, int freezedRows, int freezedColumns)` is an equivalent overload.

## Worksheet lifecycle

`Workbook.Worksheets` is the `WorksheetCollection`. `Add(name)` returns the `Worksheet`; `Add()` returns an int index. Access by name or index; remove with `RemoveAt`.

```csharp
Workbook wb = new Workbook();
Worksheet data = wb.Worksheets.Add("Data");   // returns the Worksheet
Worksheet byName = wb.Worksheets["Sheet1"];
Worksheet byIndex = wb.Worksheets[0];
wb.Worksheets.ActiveSheetIndex = data.Index;   // sheet selected on open
wb.Worksheets.RemoveAt("Sheet1");              // or RemoveAt(int)
wb.Save("output.xlsx");
```

Duplicate within a workbook with `AddCopy` (by index or name); reorder with `MoveTo`; copy across workbooks with `dest.Copy(source)`.

```csharp
Workbook wb = new Workbook("input.xlsx");
int copyIndex = wb.Worksheets.AddCopy(0);      // duplicate sheet 0
wb.Worksheets[copyIndex].MoveTo(0);            // move the copy to the front
Workbook other = new Workbook();
other.Worksheets[0].Copy(wb.Worksheets[0]);    // cross-workbook: dest.Copy(source)
other.Save("output.xlsx");
```

Tab color and visibility:

```csharp
Workbook wb = new Workbook("input.xlsx");
Worksheet sheet = wb.Worksheets[0];
sheet.TabColor = Color.Green;      // System.Drawing.Color
sheet.IsVisible = false;           // hide sheet (at least one must stay visible)
wb.Settings.ShowTabs = true;       // toggle the whole tab bar
wb.Save("output.xlsx");
```

## Rows and columns

Insert, hide, group, and size rows/columns through `Cells`. Column width is in characters; row height is in points.

```csharp
Workbook wb = new Workbook("input.xlsx");
Cells cells = wb.Worksheets[0].Cells;
cells.InsertRows(2, 3);            // 3 rows above index 2
cells.InsertColumn(1);            // 1 column at index 1
cells.HideRow(3);
cells.HideColumns(1, 2);          // hide 2 columns starting at index 1
cells.GroupRows(5, 9, false);     // outline rows 5-9, not collapsed
cells.SetColumnWidth(0, 25);      // characters
cells.SetRowHeight(0, 30);        // points
wb.Save("output.xlsx");
```

Reveal hidden rows/columns with `UnhideRow(index, height)` / `UnhideColumn(index, width)` (pass a negative width to restore the prior width). Ungroup with `UngroupRows`/`UngroupColumns`. Outline direction is `Worksheet.Outline.SummaryRowBelow` / `SummaryColumnRight`.

## Ranges

Create a range by address or by `(firstRow, firstColumn, totalRows, totalColumns)`. The range indexer is 0-based WITHIN the range. Set `Range.Name` to make it a named range; fetch with `Worksheets.GetRangeByName`.

```csharp
Workbook wb = new Workbook();
Cells cells = wb.Worksheets[0].Cells;
Aspose.Cells.Range range = cells.CreateRange(0, 0, 4, 4);  // A1:D4
range[0, 0].PutValue("top-left");   // range-local coordinates
range.Name = "MyData";              // promote to a named range
Aspose.Cells.Range named = wb.Worksheets.GetRangeByName("MyData");
wb.Save("output.xlsx");
```

Copy a range with `Copy` (all content) or `Copy(source, PasteOptions)` for paste-special:

```csharp
Workbook wb = new Workbook("input.xlsx");
Cells cells = wb.Worksheets[0].Cells;
Aspose.Cells.Range source = cells.CreateRange("A1:C5");
Aspose.Cells.Range target = cells.CreateRange("E1:G5");
target.Copy(source);                // values, formulas, formats, drawings
PasteOptions opts = new PasteOptions();
opts.PasteType = PasteType.Values;  // or Formats, Formulas, ValuesAndFormats...
target.Copy(source, opts);
wb.Save("output.xlsx");
```

Workbook-vs-worksheet-scoped defined names and the `Names` collection -> formulas-and-calculation.md. Merge/unmerge and range styling -> styles-and-formatting.md.

## Tables (ListObjects)

`Worksheet.ListObjects.Add(startRow, startCol, endRow, endCol, hasHeaders)` returns the new table's index; index into `ListObjects` to configure it.

```csharp
Workbook wb = new Workbook("input.xlsx");
Worksheet sheet = wb.Worksheets[0];
int idx = sheet.ListObjects.Add(0, 0, 9, 3, true);   // A1:D10, first row = header
ListObject table = sheet.ListObjects[idx];
table.TableStyleType = TableStyleType.TableStyleMedium9;
table.ShowTotals = true;
wb.Save("output.xlsx");
```

Convert a table back to plain cells with `table.ConvertToRange()` (or `ConvertToRange(TableToRangeOptions)`): formatting is kept, but table behavior and structured references are dropped.

## Data validation

Each `Worksheet.Validations` entry covers one or more `CellArea`s. `Add(CellArea)` returns the index; configure the returned `Validation`. For a dropdown, `InCellDropDown` must be true.

```csharp
Workbook wb = new Workbook();
Worksheet sheet = wb.Worksheets[0];
CellArea area = CellArea.CreateCellArea("A1", "A10");
Validation v = sheet.Validations[sheet.Validations.Add(area)];
v.Type = ValidationType.List;
v.InCellDropDown = true;
v.Formula1 = "Red,Green,Blue";      // inline list, or "=Sheet2!$A$1:$A$3"
wb.Save("output.xlsx");
```

Numeric range with input/error prompts:

```csharp
Workbook wb = new Workbook();
Worksheet sheet = wb.Worksheets[0];
CellArea area = CellArea.CreateCellArea("B1", "B1");
Validation v = sheet.Validations[sheet.Validations.Add(area)];
v.Type = ValidationType.WholeNumber;    // or Decimal, Date, Time, TextLength
v.Operator = OperatorType.Between;
v.Formula1 = "1";
v.Formula2 = "100";
v.AlertStyle = ValidationAlertType.Stop;   // Stop, Warning, Information
v.ErrorMessage = "Enter a whole number from 1 to 100.";
v.ShowError = true;
wb.Save("output.xlsx");
```

Add another region to an existing rule with `Validation.AddArea(cellArea)`. Types: AnyValue, WholeNumber, Decimal, List, Date, Time, TextLength, Custom.

## Protection

Worksheet protection restricts editing but does NOT encrypt data. Cells are locked by default; the classic pattern is to unlock only the cells that should stay editable, then protect. `StyleFlag.Locked = true` tells `ApplyStyle` to apply just the lock bit.

```csharp
Workbook wb = new Workbook("input.xlsx");
Worksheet sheet = wb.Worksheets[0];
Style style = wb.CreateStyle();
style.IsLocked = false;
StyleFlag flag = new StyleFlag();
flag.Locked = true;
Aspose.Cells.Range editable = sheet.Cells.CreateRange("B2", "C3");
editable.ApplyStyle(style, flag);       // unlock B2:C3
sheet.Protect(ProtectionType.All);      // lock everything else
wb.Save("output.xlsx");
```

Use `Protect(ProtectionType)` or `Protect(ProtectionType, newPassword, oldPassword)`. `ProtectionType`: All, Contents, Objects, Scenarios, Structure, Windows, None. To lock a whole row/column, use `Row.ApplyStyle` / `Column.ApplyStyle` instead of a range.

Workbook structure protection (blocks add/move/delete/rename of sheets) and file open-password encryption are separate operations:

```csharp
Workbook wb = new Workbook("input.xlsx");
wb.Protect(ProtectionType.Structure, "structPwd");   // structure lock
wb.Settings.Password = "openPwd";                    // open-password encryption
wb.Save("output.xlsx");                              // XLSX -> AES; XLS -> weaker
```

`Workbook.Settings.Password` sets the file open password (whole-file encryption); it applies to XLSX/XLS on save. Clear structure protection with `Workbook.Unprotect(password)`.

## Grid limits

XLS grids are 65536 x 256; XLSX/XLSM are 1048576 x 16384. Writing beyond the target format's limits fails or truncates - pick the format accordingly (format trade-offs -> conversion-and-rendering.md).

## Related
- styles-and-formatting.md - cell/row/column style application, merged-cell formatting, borders.
- cell-values-and-data.md - reading/writing values and iterating cells efficiently.
- formulas-and-calculation.md - defined names, named-range scope, formula recalculation.
- performance-and-large-files.md - large-sheet memory and speed, LightCells.
- conversion-and-rendering.md - format choice, grid-size limits, save options.
