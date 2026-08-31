# Styles and Formatting

Read this when formatting cells (font, color, borders, alignment, fill, number/date formats, conditional formatting, merged cells), a style change does not stick, dates display as serial numbers, or a save fails with "too many cell formats".

Core model: a cell's **value** and its **display format** are independent, and a cell's `Style` is a detached copy. Every visual change is the same shape: get or create a `Style`, mutate it, write it back. Colors are `System.Drawing.Color`.

## Pitfalls

### Style edits silently lost
Symptom: you write `cell.GetStyle().Font.IsBold = true;` (or change a color, border, or number format) and the saved file shows no change.
Cause: `Cell.GetStyle()` returns a **copy** of the cell's style, not a live reference. Mutating that copy changes nothing until you assign it back. There is no live `Cell.Style` property (it was removed in 7.1.0).
Fix: get, mutate, then `SetStyle`.

```csharp
Workbook wb = new Workbook();
Cell cell = wb.Worksheets[0].Cells["A1"];
cell.GetStyle().Font.IsBold = true;   // WRONG: mutates a discarded copy
Style s = cell.GetStyle();
s.Font.IsBold = true;
s.Font.Color = Color.Red;
cell.SetStyle(s);                     // RIGHT: persists the change
```

### Per-cell styling in loops is slow and misscaled
Symptom: formatting thousands of cells one `GetStyle`/`SetStyle` at a time is slow and bloats the file.
Cause: styling cell-by-cell allocates and pool-compares a style per cell. When a whole row, column, or range shares one look, apply it once.
Fix: build one `Style` with `Workbook.CreateStyle()`, declare a `StyleFlag` selecting which attributes to push, then apply at row/column/range scope. The `StyleFlag` is essential: only flagged attributes are written, so unrelated existing formatting on those cells is preserved.

```csharp
Workbook wb = new Workbook();
Worksheet sheet = wb.Worksheets[0];
Style st = wb.CreateStyle();
st.Font.IsBold = true;
st.ForegroundColor = Color.LightGray;
st.Pattern = BackgroundType.Solid;
st.Number = 14;
StyleFlag flag = new StyleFlag { FontBold = true, CellShading = true, NumberFormat = true };
sheet.Cells.ApplyColumnStyle(0, st, flag);   // whole column A
sheet.Cells.ApplyRowStyle(0, st, flag);      // whole row 1
```

Apply methods, each taking `(Style, StyleFlag)` plus an index where noted:
- `Cells.ApplyColumnStyle(int column, style, flag)` - one column
- `Cells.ApplyRowStyle(int row, style, flag)` - one row
- `Cells.ApplyStyle(style, flag)` - the entire worksheet
- `Range.ApplyStyle(style, flag)` - a range (below)
- `Row.ApplyStyle(style, flag)`, `Column.ApplyStyle(style, flag)` - a row/column object

Verified `StyleFlag` bools include `All`, `Font`, `FontBold`, `NumberFormat`, `Borders`, `CellShading` (fill), `HorizontalAlignment`, `WrapText`. Each style attribute has a matching flag; set `All = true` to push everything.

```csharp
Workbook wb = new Workbook();
Worksheet sheet = wb.Worksheets[0];
Style st = wb.CreateStyle();
st.HorizontalAlignment = TextAlignmentType.Center;
// Qualify: bare "Range" is ambiguous with System.Range when "using System;" is present
Aspose.Cells.Range range = sheet.Cells.CreateRange(0, 0, 1, 4);   // A1:D1
range.ApplyStyle(st, new StyleFlag { HorizontalAlignment = true });
```

### "Too many cell formats" or a corrupted file
Symptom: save throws, or Excel reports the file is corrupt / hits a format limit, after formatting many cells.
Cause: allocating a fresh `Style` per cell inflates the workbook's unique-format table. Xlsx caps unique cell formats near **64,000**; the older `.xls` binary format allows far fewer (about 4,000). Exceed it and the file is invalid.
Fix: create a small set of reusable `Style` objects and share them across cells via `SetStyle` or the bulk apply methods above. `SetStyle` pools identical styles internally, so reusing one configured object is cheap. To build a style with no `Workbook` in hand, use `CellsFactory`:

```csharp
CellsFactory factory = new CellsFactory();
Style s = factory.CreateStyle();
s.Font.IsBold = true;
```

### Dates and numbers display as raw serial numbers
Symptom: a cell shows `45658` instead of `2025-01-01`, or `0.25` instead of `25%`.
Cause: value and format are independent. Storing a `DateTime` or a fraction sets only the value; the cell keeps its `General` format and renders the underlying serial number.
Fix: put the value, then set a format on the style - a built-in id via `Style.Number`, or a pattern via `Style.Custom`.

```csharp
Workbook wb = new Workbook();
Cell cell = wb.Worksheets[0].Cells["A1"];
cell.PutValue(new DateTime(2025, 1, 1));   // value only; still shows a serial number
Style s = cell.GetStyle();
s.Number = 14;               // built-in short date; OR:
s.Custom = "yyyy-mm-dd";     // custom pattern (overrides Number)
cell.SetStyle(s);
```

`Number` and `Custom` override each other - setting one clears the other. Read a formatted date back as a typed value with `Cell.DateTimeValue` (see cell-values-and-data.md).

Useful built-in `Style.Number` ids:

| id | Type | Format |
|---:|------|--------|
| 0 | General | General |
| 1 | Decimal | 0 |
| 2 | Decimal | 0.00 |
| 4 | Decimal | #,##0.00 |
| 7 | Currency | $#,##0.00 |
| 9 | Percentage | 0% |
| 10 | Percentage | 0.00% |
| 14 | Date | short date (m/d/yyyy) |
| 15 | Date | d-mmm-yy |
| 21 | Time | h:mm:ss |
| 22 | Date-time | m/d/yy h:mm |
| 49 | Text | @ |

For anything outside this set (e.g. `yyyy-mm-dd hh:mm`, or a currency suffix), use `Style.Custom`.

### 1904 date system shifts dates by ~4 years
Symptom: dates are off by roughly four years (1462 days) when interoperating with files authored in classic Mac Excel.
Cause: workbooks use either the 1900 date system (Windows default) or the 1904 system (old Mac). The same serial number resolves to a different date under each.
Fix: match the workbook to the file's system with `Workbook.Settings.Date1904`.

```csharp
Workbook wb = new Workbook();
wb.Settings.Date1904 = true;   // interpret and store serials in the 1904 system
```

## Fonts, colors, borders, fill, alignment

All live on `Style`; set them, then `SetStyle` (or apply in bulk). Key members:
- Font: `style.Font.Name`, `.Size`, `.IsBold`, `.IsItalic`, `.Underline`, `.Color`.
- Fill: `style.ForegroundColor` is the fill color, but it only shows once `style.Pattern` is set. Use `style.Pattern = BackgroundType.Solid` for a flat fill; other `BackgroundType` values give patterns, where `style.BackgroundColor` is the second color.
- Alignment: `style.HorizontalAlignment` / `style.VerticalAlignment = TextAlignmentType.Center` (also `Left`, `Right`, `Top`, `Bottom`, `Justify`); `style.IsTextWrapped = true` to wrap.
- Borders: `style.SetBorder(BorderType, CellBorderType, Color)` per edge, or `style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin` with `.Color`. Always set line style and color together.

```csharp
Workbook wb = new Workbook();
Cell cell = wb.Worksheets[0].Cells["B2"];
Style s = cell.GetStyle();
s.Font.Name = "Calibri";
s.Font.Size = 12;
s.Font.Color = Color.White;
s.ForegroundColor = Color.SteelBlue;      // fill color...
s.Pattern = BackgroundType.Solid;         // ...needs a pattern to show
s.HorizontalAlignment = TextAlignmentType.Center;
s.IsTextWrapped = true;
s.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
cell.SetStyle(s);
```

Foreground vs background gotcha: if `Pattern` is `None`, `ForegroundColor` is ignored and reads back as `Color.Empty`; `BackgroundColor` applies only for true patterns, not for `Solid`. For a plain fill, set `ForegroundColor` + `Solid`.

## Built-in styles

`Workbook.CreateBuiltinStyle(BuiltinStyleType)` returns a ready-made named style (`Title`, `Good`, `Bad`, `Neutral`, `Note`, `Total`, plus accent and percent variants). `Workbook.GetNamedStyle(name)` fetches a named style already in the pool.

```csharp
Workbook wb = new Workbook();
Style titleStyle = wb.CreateBuiltinStyle(BuiltinStyleType.Title);
Cell cell = wb.Worksheets[0].Cells["A1"];
cell.PutValue("Report");
cell.SetStyle(titleStyle);
```

## Conditional formatting

Rules live per worksheet. Add a formatting block, attach one or more cell areas, add conditions, then style each condition. Only font (color/bold/underline/strikeout), borders, and pattern can change conditionally.

```csharp
Workbook wb = new Workbook();
Worksheet sheet = wb.Worksheets[0];
int idx = sheet.ConditionalFormattings.Add();
FormatConditionCollection fcc = sheet.ConditionalFormattings[idx];
fcc.AddArea(CellArea.CreateCellArea("A1", "A10"));
int ci = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "100", null);
// CF fills use BackgroundColor - a deliberate exception to the Foreground+Solid rule above
fcc[ci].Style.BackgroundColor = Color.Red;   // applied where A1:A10 > 100
```

- `FormatConditionType`: `CellValue`, `Expression` (formula-driven), `ColorScale`, `DataBar`, `IconSet`, and more.
- `OperatorType`: `Between`, `NotBetween`, `GreaterThan`, `LessThan`, `GreaterOrEqual`, `Equal`, and so on. For `Between` pass both `formula1` and `formula2`; for one-sided operators pass `null` as the last argument.
- For a formula rule use `FormatConditionType.Expression` and put the formula in `formula1` (e.g. `"=MOD(ROW(),2)=0"` for alternate-row shading).

Styling pivot table cells is owned by pivot-tables.md.

## Merged cells

`Cells.Merge(firstRow, firstColumn, totalRows, totalColumns)` merges; `Cells.UnMerge(...)` with the same coordinates splits. Only the top-left cell's value and format survive a merge - data in the other cells is discarded.

```csharp
Workbook wb = new Workbook();
Cells cells = wb.Worksheets[0].Cells;
cells.Merge(0, 0, 2, 3);            // merge A1:C2 (2 rows, 3 columns)
cells["A1"].PutValue("Merged title");
```

Merged regions interact with autofit and row-height calculations; that interplay is owned by worksheets-rows-columns.md.

## Related
- cell-values-and-data.md - reading typed values (DateTimeValue), PutValue, DataTable import/export
- worksheets-rows-columns.md - autofit, row height, ranges, merged-cell layout
- pivot-tables.md - styling pivot table cells
- formulas-and-calculation.md - recalculating before reading computed values
