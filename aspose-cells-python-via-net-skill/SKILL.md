---
name: aspose-cells-python
description: "Use when writing, debugging, or reviewing Python code that uses the Aspose.Cells for Python via .NET library (the `aspose-cells-python` package) to create, read, edit, convert, or render spreadsheets (xlsx, xls, xlsm, csv, tsv, ods, and spreadsheet-to-PDF/HTML/image). Trigger for any Aspose.Cells task: applying a license or fixing the evaluation watermark, setting cell values and number formats, styles and conditional formatting, formulas and recalculation, charts, pivot tables, inserting images, shapes, or comments, smart-markers reporting, importing or exporting data, converting Excel to PDF, large-file memory tuning, or Linux/Docker deployment. Use it whenever you see the `aspose.cells` import, a `Workbook` or `Worksheet` object, or an `aspose-cells-python` dependency. Do NOT use for Aspose.Cells for .NET (C#), Java, Node.js, or C++; for other Aspose products (Words, PDF, Slides, Email, Imaging); for Aspose.Cells GridJs or GridWeb UI components."
license: Proprietary. See Aspose EULA.
---

# Aspose.Cells for Python via .NET

## Mental model

Aspose.Cells is an in-memory **document object model**, not a running copy of Excel.
Nothing happens automatically: formulas do not recalculate, pivot tables and charts do
not refresh, display formats are not inferred from values, and style edits do not write
themselves back. Each of those is an explicit API call, and most real-world bugs are a
missing one. All row/column/sheet indices are **0-based** (row 0 = Excel row 1).

The package is imported as `import aspose.cells as gc`. Many types live in submodules
(`gc.charts`, `gc.pivot`, `gc.slicers`, `gc.drawing`, `gc.rendering`,
`gc.digitalsignatures`, `gc.tables`, `gc.utility`, `gc.vba`, ...). Colors come from
`aspose.pydrawing` (`import aspose.pydrawing as drawing`). Method and property names use
`snake_case` (e.g. `put_value`, `set_license`, `calculate_formula`, `is_licensed`), and
there is no `cells["A1"]` indexer - use `cells.get("A1")` or `cells.get(row, col)`.

## Golden rules

Each rule links to the reference file that owns the full pattern; read it before
writing code in that area.

1. **License first.** Set it once at process startup, before any `Workbook` is created;
   verify with `workbook.is_licensed`. A wheel released after your subscription's end date
   silently reverts to evaluation mode - the usual cause of "watermark appeared after an
   upgrade". -> `references/setup-and-licensing.md`
2. **Nothing recalculates itself.** After adding or editing formulas, and before reading
   computed values, call `workbook.calculate_formula()`. -> `references/formulas-and-calculation.md`
3. **Value != format.** A `datetime` or number displays as a raw serial number until you
   set a date/number format on the cell's style. -> `references/styles-and-formatting.md`
4. **`get_style()` returns a copy.** Mutate it, then write back with `set_style()`;
   bulk-format with `workbook.create_style()` + `StyleFlag` on a range, row, or column.
   -> `references/styles-and-formatting.md`
5. **Reuse styles.** Never create a new `Style` per cell in a loop - xlsx caps unique
   cell formats at ~64,000. -> `references/styles-and-formatting.md`
6. **Don't scan by index.** `cells.get(r, c)` instantiates empty cells; iterate
   `for row in cells.rows:` and then the cells in each row.
   -> `references/cell-values-and-data.md`
7. **Everything is 0-based**, and `insert_rows`/`delete_rows` auto-shift formulas and
   references - don't adjust them again yourself; delete bottom-up in loops.
   -> `references/worksheets-rows-columns.md`
8. **State PDF layout explicitly**: `PdfSaveOptions.one_page_per_sheet`,
   `all_columns_in_one_page_per_sheet`, or `PageSetup.fit_to_pages_wide/tall`.
   -> `references/conversion-and-rendering.md`
9. **Streams**: after `save(stream, ...)`, reset `stream.seek(0)` before returning
   it, and always pass an explicit `SaveFormat`. -> `references/conversion-and-rendering.md`
10. **Charts and pivots start empty**: call `chart.calculate()` before `to_image`/`to_pdf`;
   populate a pivot with `calculate_data(PivotTableCalculateOption(refresh_data=True))`
   or `pt.pivot_cache.refresh(); pt.calculate_data()`
   (`PivotTable.refresh_data()` is obsolete in 26.7 - use `PivotCache.refresh()`).
   -> `references/charts.md`, `references/pivot-tables.md`
11. **Big files**: open with `MemorySetting.MEMORY_PREFERENCE` or stream with the
   LightCells API. -> `references/performance-and-large-files.md`
12. **One `Workbook` per thread** - the object model is not thread-safe.
   -> `references/performance-and-large-files.md`
13. **Linux/.NET 6+ needs setup**: Aspose.Cells >= 22.10.1 renders via SkiaSharp there -
   the Python package bundles it; install `libfontconfig1` and ship fonts.
   -> `references/cross-platform-deployment.md`
14. **Missing fonts shift layout silently** in PDF/image output - configure
   `FontConfigs` and set `default_font` in save/render options.
   -> `references/conversion-and-rendering.md`
15. **Macros need a macro format**: saving an `.xlsm`/`.xls` with VBA to `.xlsx`
   silently drops all macros. `Workbook.vba_project` is never null on a new workbook;
   add modules to it (`VbaProject` has no public constructor). Sign *last*, after all
   edits. -> `references/vba-macros.md`
16. **Signatures cover file bytes**: any edit or re-save after signing invalidates the
   signature. Sign last, right before the final save. -> `references/digital-signatures-and-document-properties.md`
17. **Slicers attach to a pivot or ListObject**, not a raw cell range; a slicer over a
   pivot only filters the cache, so refresh the pivot data before reading output.
   Sparkline colors are `CellsColor`, not `Color`. -> `references/slicers-and-sparklines.md`

## Quickstart

```python
import aspose.cells as gc

# Once per process, before any Workbook - otherwise output carries an evaluation watermark.
gc.License().set_license("Aspose.Cells.lic")

workbook = gc.Workbook()                 # or gc.Workbook("input.xlsx")
sheet = workbook.worksheets[0]
sheet.cells.get("A1").put_value("Price")
sheet.cells.get("B1").put_value(42.5)
sheet.cells.get("C1").formula = "=B1*2"
workbook.calculate_formula()             # formulas never calculate themselves
workbook.save("output.xlsx", gc.SaveFormat.XLSX)
```

## Version assumptions

Written for **Aspose.Cells for Python via .NET 26.7.0** (the `aspose-cells-python` wheel) on
Python 3.8+. Every class, enum, and submodule name shown in SKILL.md and `references/` was
checked by introspection against that exact package, and the most error-prone method calls
(such as `License.set_license`, `Cells.import_array`, `Cell.put_value`, `Workbook.calculate_formula`,
`VbaProject.sign`) were introspected too. Graphics on non-Windows uses the bundled SkiaSharp
backend since 22.10.1. API surface follows https://reference.aspose.com/cells/pythonnet/.
If you pin a different version, re-run an introspection check before trusting a snippet here.

## When to read which reference

| Task or symptom                                                                                 | Read                                        |
| ----------------------------------------------------------------------------------------------- | ------------------------------------------- |
| Install, apply/verify a license, evaluation watermark or truncated output                       | `references/setup-and-licensing.md`         |
| Works on Windows; fails or renders wrong fonts on Linux/Docker/macOS                            | `references/cross-platform-deployment.md`   |
| Read/write cell values, wrong types on read, import/export lists/objects                        | `references/cell-values-and-data.md`        |
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
| Apply license | `gc.License().set_license("Aspose.Cells.lic")` before any `Workbook` | setup-and-licensing |
| Check license | `workbook.is_licensed` | setup-and-licensing |
| Write a value | `cell.put_value(42.5)` / `put_value(datetime)` / `put_value("text")` | cell-values-and-data |
| Read a value safely | branch on `cell.type`, then `cell.double_value`/`string_value`/etc. | cell-values-and-data |
| Date/number display | `style.number = 14` (m/d/yyyy) on the cell style | styles-and-formatting |
| Edit a style | `s = cell.get_style(); s.font.is_bold = True; cell.set_style(s)` | styles-and-formatting |
| Recalculate formulas | `workbook.calculate_formula()` after edits, before reads | formulas-and-calculation |
| Find cells | `cells.find(value, prev, opts)` loop; `sheet.replace(a, b)` | cell-values-and-data |
| Iterate only real cells | `for row in cells.rows: for cell in row:` | cell-values-and-data |
| Insert/delete rows | `sheet.cells.insert_rows(row, count)`; delete bottom-up in loops | worksheets-rows-columns |
| Data validation dropdown | `v = ...; v.type = gc.ValidationType.LIST; v.formula1 = "a,b,c"` | worksheets-rows-columns |
| Protect a sheet | `sheet.protect(gc.ProtectionType.CONTENTS)`; unlock cells first | worksheets-rows-columns |
| Embed a file | `sheet.ole_objects.add(row, col, h, w, icon, ...)` | ole-objects |
| Add a comment | `sheet.comments.add("A1")` then set text/author | images-and-shapes |
| Add a picture | `sheet.pictures.add(row, col, bytes)` | images-and-shapes |
| Add a chart | `idx = sheet.charts.add(gc.charts.ChartType.COLUMN, r, c, r2, c2)` | charts |
| Render a chart | `chart.calculate(); chart.to_image(stream, ImageOrPrintOptions)` | charts |
| Create a pivot | `pt = sheet.pivot_tables.add(...); pt.calculate_data(...)` | pivot-tables |
| Refresh a pivot | `pt.calculate_data(gc.pivot.PivotTableCalculateOption(refresh_data=True))` | pivot-tables |
| Slicer over a table | `sheet.slicers.add(list_object, column_index, "E2")` | slicers-and-sparklines |
| Add a sparkline | `sheet.sparkline_groups.add(gc.charts.SparklineType.LINE, "A1:B5", True, location)` | slicers-and-sparklines |
| Add a macro module | `vba.modules.add(gc.vba.VbaModuleType.PROCEDURAL, "Name"); module.codes = "..."` | vba-macros |
| Sign a workbook | `wb.add_digital_signature(sigs)` last, before final save | digital-signatures-and-document-properties |
| Custom property | `wb.custom_document_properties.add("Key", value)`; read via `.value` | digital-signatures-and-document-properties |
| To PDF, 1 sheet/page | `SaveFormat.PDF` + `PdfSaveOptions(one_page_per_sheet=True)` | conversion-and-rendering |
| Reset a stream | `save(stream, SaveFormat.XLSX); stream.seek(0)` | conversion-and-rendering |
| Smart markers fill | `d = gc.WorkbookDesigner(book); d.set_data_source(tbl); d.process(); d.workbook.save(...)` | smart-markers-reporting |
| Large file, low memory | `LoadOptions(memory_setting=gc.MemorySetting.MEMORY_PREFERENCE)` | performance-and-large-files |

## Verifying the docs

Two verifiers guard every Python snippet and inline reference in this skill. They
both run against the installed `aspose-cells-python` package; exit code 0 means clean.

```powershell
# in the skill repo root
# 1) Syntax-compile every ```python block (catches bad indentation,
#    unbalanced braces, syntax errors - like a linter for the snippets).
python verifier/verify.py --compile

# 2) Reflection scan of inline references (checks that every type/method/enum
#    member named in prose and tables actually exists in the installed package).
python verifier/verify.py
```

The compile run prints `COMPILE OK: N   ERRORS: M   BLOCKS: K`. Any `ERRORS`
line is a real doc bug. All blocks are compiled independently; a `class Product`
declared in one block is NOT visible to another, so keep each snippet
self-contained (or redeclare the helper in the block that uses it). A `var wb`/
`sheet`/`cells` referenced without a declaration in the .NET skill is auto-injected
here via explicit `wb = gc.Workbook()` / `sheet = wb.worksheets[0]` lines in each
snippet.

The reflection run prints `OK: n   NOT FOUND: m   SKIPPED: k`. Exit code 0 means
no NOT FOUND. A `NOT FOUND` line is a real doc bug - either fix the doc or, if the
API is genuinely gone, update the reference. `SKIPPED` counts expressions the
scanner cannot prove (instance-only attributes, loop variables) - review them by
eye, not by script. The scanner resolves variable declarations (`wb = gc.Workbook()`,
`sheet = wb.worksheets[0]`) within a code block, so most snippets get full chain
validation. When upgrading the package version, bump it in SKILL.md and re-run both
verifiers.
