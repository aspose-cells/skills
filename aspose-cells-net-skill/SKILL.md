---
name: aspose-cells-net
description: "Use when writing, debugging, or reviewing C#/.NET code that uses the Aspose.Cells for .NET library to create, read, edit, convert, or render spreadsheets (xlsx, xls, xlsm, csv, tsv, ods, and spreadsheet-to-PDF/HTML/image). Trigger for any Aspose.Cells task: applying a license or fixing the evaluation watermark, setting cell values and number formats, styles and conditional formatting, formulas and recalculation, charts, pivot tables, inserting images, shapes, or comments, smart-markers reporting, importing or exporting DataTables, converting Excel to PDF, large-file memory tuning, or Linux/.NET6/Docker deployment. Use it whenever you see the Aspose.Cells namespace, a Workbook or Worksheet object, or an Aspose.Cells NuGet reference. Do NOT use for Aspose.Cells for Java, Python, Node.js, or C++; for other Aspose products (Words, PDF, Slides, Email, Imaging); for Aspose.Cells GridJs or GridWeb UI components."
license: Proprietary. See Aspose EULA.
---

# Aspose.Cells for .NET

## Mental model

Aspose.Cells is an in-memory **document object model**, not a running copy of Excel.
Nothing happens automatically: formulas do not recalculate, pivot tables and charts do
not refresh, display formats are not inferred from values, and style edits do not write
themselves back. Each of those is an explicit API call, and most real-world bugs are a
missing one. All row/column/sheet indices are **0-based** (row 0 = Excel row 1).

## Golden rules

Each rule links to the reference file that owns the full pattern; read it before
writing code in that area.

1. **License first.** Set it once at process startup, before any `Workbook` is created;
   verify with `Workbook.IsLicensed`. A DLL released after your subscription's end date
   silently reverts to evaluation mode - the usual cause of "watermark appeared after a
   NuGet upgrade". -> `references/setup-and-licensing.md`
2. **Nothing recalculates itself.** After adding or editing formulas, and before reading
   computed values, call `workbook.CalculateFormula()`. -> `references/formulas-and-calculation.md`
3. **Value != format.** A `DateTime` or number displays as a raw serial number until you
   set a date/number format on the cell's style. -> `references/styles-and-formatting.md`
4. **`GetStyle()` returns a copy.** Mutate it, then write back with `SetStyle()`;
   bulk-format with `workbook.CreateStyle()` + `StyleFlag` on a range, row, or column.
   -> `references/styles-and-formatting.md`
5. **Reuse styles.** Never create a new `Style` per cell in a loop - xlsx caps unique
   cell formats at ~64,000. -> `references/styles-and-formatting.md`
6. **Don't scan by index.** `cells[r, c]` instantiates empty cells; iterate
   `foreach (Row row in cells.Rows)` and then the cells in each row.
   -> `references/cell-values-and-data.md`
7. **Everything is 0-based**, and `InsertRows`/`DeleteRows` auto-shift formulas and
   references - don't adjust them again yourself; delete bottom-up in loops.
   -> `references/worksheets-rows-columns.md`
8. **State PDF layout explicitly**: `PdfSaveOptions.OnePagePerSheet`,
   `AllColumnsInOnePagePerSheet`, or `PageSetup.FitToPagesWide/Tall`.
   -> `references/conversion-and-rendering.md`
9. **Streams**: after `Save(stream, ...)`, reset `stream.Position = 0` before returning
   it, and always pass an explicit `SaveFormat`. -> `references/conversion-and-rendering.md`
10. **Charts and pivots start empty**: call `chart.Calculate()` before `ToImage`/`ToPdf`;
    populate a pivot with `CalculateData(new PivotTableCalculateOption { RefreshData = true })`
    or `pt.PivotCache.Refresh(); pt.CalculateData();` (`PivotTable.RefreshData()` is obsolete
    in 26.7 - use `PivotCache.Refresh()`).
    -> `references/charts.md`, `references/pivot-tables.md`
11. **Big files**: open with `MemorySetting.MemoryPreference` or stream with the
    LightCells API. -> `references/performance-and-large-files.md`
12. **One `Workbook` per thread** - the object model is not thread-safe.
    -> `references/performance-and-large-files.md`
13. **Linux/.NET 6+ needs setup**: Aspose.Cells >= 22.10.1 renders via SkiaSharp there -
    add the matching `SkiaSharp.NativeAssets.Linux`, install `libfontconfig1`, ship
    fonts. -> `references/cross-platform-deployment.md`
14. **Missing fonts shift layout silently** in PDF/image output - configure
    `FontConfigs` and set `DefaultFont` in save/render options.
    -> `references/conversion-and-rendering.md`
15. **Macros need a macro format**: saving an `.xlsm`/`.xls` with VBA to `.xlsx`
    silently drops all macros. `Workbook.VbaProject` is never null on a new workbook;
    add modules to it (`VbaProject` has no public constructor). Sign *last*, after all
    edits. -> `references/vba-macros.md`
16. **Signatures cover file bytes**: any edit or re-save after signing invalidates the
    signature. Sign last, right before the final save. -> `references/digital-signatures-and-document-properties.md`
17. **Slicers attach to a pivot or ListObject**, not a raw cell range; a slicer over a
    pivot only filters the cache, so refresh the pivot data before reading output.
    Sparkline colors are `CellsColor`, not `Color`. -> `references/slicers-and-sparklines.md`

## Quickstart

```csharp
using Aspose.Cells;   // dotnet add package Aspose.Cells

// Once per process, before any Workbook - otherwise output carries an evaluation watermark.
new License().SetLicense("Aspose.Cells.lic");

var workbook = new Workbook();                 // or new Workbook("input.xlsx")
Worksheet sheet = workbook.Worksheets[0];
sheet.Cells["A1"].PutValue("Price");
sheet.Cells["B1"].PutValue(42.5);
sheet.Cells["C1"].Formula = "=B1*2";
workbook.CalculateFormula();                   // formulas never calculate themselves
workbook.Save("output.xlsx", SaveFormat.Xlsx);
```

## Version assumptions

Written for **Aspose.Cells for .NET 26.7.0** on .NET 8+, and every API shown in
SKILL.md and `references/` is verified by reflection against that exact package (see
"Verifying the docs" below). Graphics: `System.Drawing.Common` on Windows-targeted
builds; SkiaSharp on cross-platform .NET 6+ builds since 22.10.1. API surface follows
https://reference.aspose.com/cells/net/. If you pin a different major version, re-run
the verifier before trusting a code snippet here.

## When to read which reference

| Task or symptom                                                                                 | Read                                        |
| ----------------------------------------------------------------------------------------------- | ------------------------------------------- |
| Install, apply/verify a license, evaluation watermark or truncated output                       | `references/setup-and-licensing.md`         |
| Works on Windows; fails or renders wrong fonts on Linux/Docker/macOS/.NET 6+                    | `references/cross-platform-deployment.md`   |
| Read/write cell values, wrong types on read, import/export DataTable/arrays/objects             | `references/cell-values-and-data.md`        |
| Fonts, colors, borders, number/date formats, conditional formatting, formats that "don't stick" | `references/styles-and-formatting.md`       |
| Formulas return null/0/stale, slow calculation, circular references, custom functions           | `references/formulas-and-calculation.md`    |
| Add/copy/delete sheets, insert/delete rows or columns, ranges, tables, validation, protection   | `references/worksheets-rows-columns.md`     |
| Insert or position images, shapes, cell comments                                                | `references/images-and-shapes.md`           |
| Create or customize charts; chart image blank or wrong                                          | `references/charts.md`                      |
| Create or refresh pivot tables; pivot shows no or stale data                                    | `references/pivot-tables.md`                |
| Save formats, Excel to PDF/HTML/image, streaming downloads, print setup                         | `references/conversion-and-rendering.md` |
| OutOfMemory, huge files, slow bulk IO, threading, timeouts                                      | `references/performance-and-large-files.md` |
| Fill an Excel template from data with &= smart markers                                          | `references/smart-markers-reporting.md` |
| Add/edit/read VBA macros; macros disappear on save; sign a macro project                         | `references/vba-macros.md` |
| Sign a workbook, verify signatures, built-in/custom document properties                          | `references/digital-signatures-and-document-properties.md` |
| Interactive filters (slicers) over pivots/tables; mini sparkline charts                          | `references/slicers-and-sparklines.md` |
| Embed another file (Word/PDF/Excel/image) inside a worksheet cell                                | `references/ole-objects.md` |

## Cheat sheet

A fast index of the most-reached-for calls. Each row resolves to one line in the
referenced file for the full pattern and pitfalls.

| Need | Code | Reference |
|---|---|---|
| Apply license | `new License().SetLicense("Aspose.Cells.lic")` before any `Workbook` | setup-and-licensing |
| Check license | `workbook.IsLicensed` | setup-and-licensing |
| Write a value | `cell.PutValue(42.5)` / `PutValue(DateTime)` / `PutValue("text")` | cell-values-and-data |
| Read a value safely | branch on `cell.Type`, then `cell.DoubleValue`/`StringValue`/etc. | cell-values-and-data |
| Date/number display | `style.Number = 14` (m/d/yyyy) on the cell style | styles-and-formatting |
| Edit a style | `var s = cell.GetStyle(); s.Font.Bold = true; cell.SetStyle(s);` | styles-and-formatting |
| Recalculate formulas | `workbook.CalculateFormula()` after edits, before reads | formulas-and-calculation |
| Find cells | `cells.Find(value, prev, opts)` loop; `sheet.Replace(a, b)` | cell-values-and-data |
| Iterate only real cells | `foreach (Row r in cells.Rows) foreach (Cell c in r)` | cell-values-and-data |
| Insert/delete rows | `sheet.Cells.InsertRows(row, count)`; delete bottom-up in loops | worksheets-rows-columns |
| Data validation dropdown | `Validation v = ...; v.Type = ValidationType.List; v.Formula1 = "a,b,c";` | worksheets-rows-columns |
| Protect a sheet | `sheet.Protect(ProtectionType.Contents)`; unlock cells first | worksheets-rows-columns |
| Embed a file | `sheet.OleObjects.Add(row, col, bytes, "pdf")` | ole-objects |
| Add a comment | `sheet.Comments.Add(cellName)` then set text/author | images-and-shapes |
| Add a picture | `sheet.Pictures.Add(row, col, bytes)` | images-and-shapes |
| Add a chart | `var idx = sheet.Charts.Add(ChartType.Column, row, col, row2, col2)` | charts |
| Render a chart | `chart.Calculate(); chart.ToImage(stream, ImageOrPrintOptions)` | charts |
| Create a pivot | `pt = sheet.PivotTables.Add(...);` then `pt.CalculateData(...)` | pivot-tables |
| Refresh a pivot | `pt.CalculateData(new PivotTableCalculateOption { RefreshData = true })` | pivot-tables |
| Slicer over a table | `sheet.Slicers.Add(listObject, columnIndex, "E2")` | slicers-and-sparklines |
| Add a sparkline | `sheet.SparklineGroups.Add(SparklineType.Line, "A1:B5", true, location)` | slicers-and-sparklines |
| Add a macro module | `vba.Modules.Add(VbaModuleType.Procedural, "Name"); module.Codes = "...";` | vba-macros |
| Sign a workbook | `wb.AddDigitalSignature(sigs)` last, before final save | digital-signatures-and-document-properties |
| Custom property | `wb.CustomDocumentProperties.Add("Key", value)`; read via `.Value` | digital-signatures-and-document-properties |
| To PDF, 1 sheet/page | `SaveFormat.Pdf` + `PdfSaveOptions { OnePagePerSheet = true }` | conversion-and-rendering |
| Reset a stream | `Save(stream, SaveFormat.Xlsx); stream.Position = 0;` | conversion-and-rendering |
| Smart markers fill | `var d = new WorkbookDesigner(book); d.SetDataSource(tbl); d.Process(); d.Workbook.Save(...)` | smart-markers-reporting |
| Large file, low memory | `LoadOptions { MemorySetting = MemorySetting.MemoryPreference }` | performance-and-large-files |

## Verifying the docs

Two verifiers guard every C# snippet and inline reference in this skill. They
both run against the pinned Aspose.Cells package; exit code 0 means clean.

```powershell
# in the skill repo root
# 1) Compile every ```csharp block with Roslyn (catches wrong overloads,
#    bad argument types, misspelled members - like javac for the Java skill).
dotnet run --project verifier/AsposeCellsVerifier.csproj -c Release -- --compile

# 2) Reflection scan of inline references (checks that every type/member
#    named in prose and tables actually exists in 26.7).
dotnet run --project verifier/AsposeCellsVerifier.csproj -c Release
```

The compile run prints `COMPILE OK: N   ERRORS: M   BLOCKS: K`. Any `ERRORS`
line is a real doc bug - the report maps each error back to its source .md and
line. All blocks are compiled into ONE Roslyn compilation unit, so a type
declared in one block (e.g. `class Product`) is visible to later blocks that
use it. A `var wb`/`sheet`/`cells` referenced without a declaration is
auto-injected, mirroring the Java skill's extract.py wrapper.

The reflection run prints `OK: n NOT FOUND: m SKIPPED: k`. Exit code 0 means
no NOT FOUND. A `NOT FOUND` line is a real doc bug - either fix the doc or, if
the API is genuinely gone, update the reference. `SKIPPED` counts expressions
the scanner cannot prove (loop variables, `Console`, `typeof`, multi-line
chains) - review them by eye, not by script. The verifier resolves variable
declarations (`Workbook wb = ...`, `var sheet = ...`) across a code block, so
most snippets get full chain validation. When upgrading the package version,
bump it in `verifier/AsposeCellsVerifier.csproj` and re-run both verifiers.
