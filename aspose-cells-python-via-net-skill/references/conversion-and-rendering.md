# Conversion, Rendering, and Saving

Read when saving to any format, converting Excel to PDF/HTML/image/JSON/CSV, controlling PDF page layout, a downloaded file arrives corrupt, or choosing xls vs xlsx.

Conversion means load a workbook in one format and `save` it in another. Rendering (PDF/image/XPS/HTML) is layout-sensitive: output follows each worksheet's `page_setup` unless you override it. PDF, XPS, and image formats are save-only render targets - they cannot be loaded back into a `Workbook`.

## Pitfalls

### Excel-to-PDF splits pages wrong or content overflows
- Symptom: columns spill onto extra pages, a wide table breaks mid-row, or one sheet becomes dozens of PDF pages.
- Cause: paging defaults to each sheet's `page_setup` (paper size, margins, manual breaks). Fit-to-page is opt-in, not automatic.
- Fix: force layout via `PdfSaveOptions` or per-sheet `page_setup`:
  - `PdfSaveOptions.one_page_per_sheet = True` - the entire sheet on one page.
  - `PdfSaveOptions.all_columns_in_one_page_per_sheet = True` - all columns one page wide, rows may overflow.
  - `sheet.page_setup.fit_to_pages_wide = 1; fit_to_pages_tall = 0` - scale to 1 page wide, unlimited tall.
  - Constrain output with `page_setup.print_area`, `orientation`, `paper_size`. See the PDF how-to below.

### Missing fonts substitute silently and shift PDF/image layout
- Symptom: rendered text differs from Excel - wrong glyphs, changed widths, misaligned columns. No exception is thrown.
- Cause: the render font must exist on the machine. If absent, Aspose.Cells substitutes the closest available font (finally Arial), which changes metrics.
- Fix: register the fonts and set a deterministic fallback:
  - `PdfSaveOptions.default_font` / `ImageOrPrintOptions.default_font` - fallback face for unmapped glyphs.
  - `PdfSaveOptions.is_font_substitution_char_granularity = True` - substitute per character, not per run.
  - `FontConfigs.set_font_folder(dir, True)` at startup, before any render. Cross-platform/Docker font setup -> cross-platform-deployment.md.
  - Detect substitutions with an `IWarningCallback` (below) instead of eyeballing output.

### Downloaded xlsx/pdf arrives corrupt or zero-length (web apps)
- Symptom: the browser downloads a file that will not open; bytes are truncated or empty.
- Cause: after `save(stream, ...)` the stream position sits at the end, so code that copies "from the current position" sends nothing; or the declared `SaveFormat`, MIME type, and extension disagree.
- Fix: reset `ms.seek(0)` before reading (or use `ms.getvalue()`), and keep format/MIME/extension consistent. See the web-download how-to.

### SaveFormat.EXCEL_97_TO_2003 (xls) truncates or fails on large data
- Symptom: rows past 65,536 or columns past 256 vanish, or Save fails, when writing xls.
- Cause: the legacy .xls (BIFF8) grid is capped at 65,536 rows x 256 columns; xlsx/xlsb allow 1,048,576 x 16,384.
- Fix: save large data as `SaveFormat.XLSX` (or `XLSB` for size and speed). Use `EXCEL_97_TO_2003` only when a consumer truly requires .xls, and keep data within the legacy limits.

### PDF hairlines look thicker in Adobe Reader
- Symptom: borders/gridlines render heavier in Adobe Reader than in Excel or other viewers.
- Cause: Adobe Reader's "Enhance thin lines" / "Smooth line art" display settings, not an output defect. The PDF geometry is correct.
- Fix: nothing in code. In Adobe Reader: Edit > Preferences > Page Display, uncheck "Smooth line art" and "Enhance thin lines". Other viewers are unaffected.

## Save formats and Save overloads

`Workbook.save` picks the target from the `SaveFormat` enum (or a `*SaveOptions` object).

| Target | SaveFormat member | Notes |
| :- | :- | :- |
| xlsx | `XLSX` | default modern workbook |
| xlsm | `XLSM` | macro-enabled |
| xlsb | `XLSB` | binary; smallest and fastest for big data |
| xls | `EXCEL_97_TO_2003` | legacy; 65,536 x 256 limit |
| csv / tsv | `CSV` / `TSV` | active sheet only by default |
| ods | `ODS` | OpenDocument Calc |
| pdf | `PDF` | pair with `PdfSaveOptions` |
| html / mht | `HTML` / `MHTML` | pair with `HtmlSaveOptions` |
| json | `JSON` | pair with `JsonSaveOptions` |

`save(path)` infers the format from the file extension; pass a `SaveFormat` explicitly when the extension is missing or ambiguous.

```python
import aspose.cells as gc

wb = gc.Workbook("book.xlsx")
wb.save("out.pdf")                         # format inferred from ".pdf"
wb.save("out.data", gc.SaveFormat.XLSX)    # explicit format
with open("copy.xlsx", "wb") as f:
    wb.save(f, gc.SaveFormat.XLSX)         # to a stream
```

Also `save(path, SaveOptions)` and `save(Stream, SaveOptions)` for fine-grained control.

## Excel to PDF

Set options on `PdfSaveOptions`. If the sheet has formulas, call `calculate_formula()` (or set `PdfSaveOptions.calculate_formula = True`) before saving so rendered values are current.

```python
import aspose.cells as gc

wb = gc.Workbook("book.xlsx")
wb.calculate_formula()                       # refresh formula results first
opts = gc.PdfSaveOptions()
opts.one_page_per_sheet = True               # whole sheet -> one page
opts.compliance = gc.rendering.PdfCompliance.PDF_A1B  # archival PDF/A-1b (else PDF 1.7)
wb.save("out.pdf", opts)
```

Control page layout per sheet through `page_setup`:

```python
import aspose.cells as gc

wb = gc.Workbook("book.xlsx")
ps = wb.worksheets[0].page_setup
ps.fit_to_pages_wide = 1                        # all columns across 1 page
ps.fit_to_pages_tall = 0                        # 0 = as many row-pages as needed
ps.print_area = "A1:H200"
ps.orientation = gc.PageOrientationType.LANDSCAPE
ps.paper_size = gc.PaperSizeType.PAPER_A4
wb.save("out.pdf", gc.SaveFormat.PDF)
```

Render only chosen sheets with `PdfSaveOptions.sheet_set`:

```python
import aspose.cells as gc
import aspose.cells.rendering as rd

wb = gc.Workbook("book.xlsx")
opts = gc.PdfSaveOptions()
opts.sheet_set = rd.SheetSet.active              # or SheetSet.all (includes hidden)
# opts.sheet_set = rd.SheetSet([0, 2])           # sheets 0 and 2, in this order
wb.save("selected.pdf", opts)
```

`PdfSaveOptions` also exposes `security_options` (encryption/permissions), `producer`, and `created_time`.

### Detect font substitution while rendering

Implement `IWarningCallback` and assign it to the save options' `warning_callback`:

```python
import aspose.cells as gc

class FontWarner(gc.IWarningCallback):
    def warning(self, info):
        if info.warning_type == gc.WarningType.FONT_SUBSTITUTION:
            print(info.description)   # logs the missing font + cell

wb = gc.Workbook("book.xlsx")
opts = gc.PdfSaveOptions()
opts.warning_callback = FontWarner()
wb.save("out.pdf", opts)
```

## Excel to image (PNG/JPEG/TIFF)

Use `SheetRender` (one worksheet) or `WorkbookRender` (whole book), configured with `ImageOrPrintOptions`. `ImageType` defaults to PNG.

```python
import aspose.cells as gc
import aspose.cells.rendering as rd

wb = gc.Workbook("book.xlsx")
opts = rd.ImageOrPrintOptions()
opts.image_type = gc.drawing.ImageType.PNG
opts.horizontal_resolution = 200              # DPI; raise for print quality
opts.vertical_resolution = 200
opts.one_page_per_sheet = True
sr = rd.SheetRender(wb.worksheets[0], opts)
for p in range(sr.page_count):
    sr.to_image(p, f"page{p}.png")            # one file per page
```

Whole workbook to a single multi-page TIFF:

```python
import aspose.cells as gc
import aspose.cells.rendering as rd

wb = gc.Workbook("book.xlsx")
opts = rd.ImageOrPrintOptions()
opts.image_type = gc.drawing.ImageType.TIFF
wr = rd.WorkbookRender(wb, opts)
wr.to_image("workbook.tiff")                  # all pages in one TIFF
```

`SheetRender.to_tiff(path)` writes a multi-page TIFF for a single sheet. `to_image` overloads also accept a `Stream`.

## Excel to HTML

```python
import aspose.cells as gc

wb = gc.Workbook("book.xlsx")
opts = gc.HtmlSaveOptions()
opts.export_images_as_base64 = True             # self-contained single .html
opts.export_active_worksheet_only = True       # omit to export every sheet
wb.save("out.html", opts)
```

`HtmlSaveOptions.image_options` (an `ImageOrPrintOptions`) tunes how shapes/charts rasterize; `export_grid_lines` and `page_title` are also available.

## Excel to and from JSON

Export the active sheet with `SaveFormat.JSON`, or a range with `JsonUtility`:

```python
import aspose.cells as gc
import aspose.cells.utility as ut

wb = gc.Workbook("book.xlsx")
wb.save("out.json", gc.SaveFormat.JSON)         # active sheet
rng = wb.worksheets[0].cells.create_range("A1:C10")
json = ut.JsonUtility.export_range_to_json(rng, gc.JsonSaveOptions())
```

Import a JSON string into cells with a layout, or load a whole `.json` file:

```python
import aspose.cells as gc
import aspose.cells.utility as ut
import json

wb = gc.Workbook()
cells = wb.worksheets[0].cells
with open("data.json", "r", encoding="utf-8") as f:
    data = f.read()
ut.JsonUtility.import_data(data, cells, 0, 0, ut.JsonLayoutOptions())
whole = gc.Workbook("data.json", gc.LoadOptions(gc.LoadFormat.JSON))
```

## Excel to CSV/TSV/text

CSV/TSV/TXT save the active sheet only unless told otherwise. Use `TxtSaveOptions` for delimiter and multi-sheet control.

```python
import aspose.cells as gc

wb = gc.Workbook("book.xlsx")
wb.save("out.csv", gc.SaveFormat.CSV)           # active sheet, comma-separated
txt = gc.TxtSaveOptions(gc.SaveFormat.CSV)
txt.separator = ';'                             # custom delimiter
txt.export_all_sheets = True                    # concatenate all sheets
wb.save("all.csv", txt)
```

## LowCode one-call converters

The `aspose.cells.lowcode` module runs whole-file conversions without building a `Workbook`.

```python
import aspose.cells.lowcode as lc

lc.PdfConverter.process("book.xlsx", "out.pdf")
lc.SpreadsheetConverter.process("book.xlsx", "out.xlsb")
lc.SpreadsheetLocker.process("book.xlsx", "locked.xlsx", "pw", "pw")
```

For finer control pass `LowCodeLoadOptions` + `LowCodeSaveOptions`, e.g. `PdfConverter.process(load, save)`. Siblings: `HtmlConverter`, `ImageConverter`, `JsonConverter`, `TextConverter`, `SpreadsheetMerger.process(inputs, output)`, and `SpreadsheetSplitter.process(input, output)`.

## Web download (framework-neutral)

Save to a `BytesIO`, rewind, then hand the bytes to the response. MIME types: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` for xlsx, `application/vnd.ms-excel` for xls, `application/pdf` for pdf.

```python
import io
import aspose.cells as gc

wb = gc.Workbook("book.xlsx")
ms = io.BytesIO()
wb.save(ms, gc.SaveFormat.XLSX)                 # explicit format; match MIME + name
ms.seek(0)                                     # save leaves position at the end
data = ms.getvalue()                            # the file bytes
```

## Page setup essentials

Header/footer scripts use `&P` (page), `&N` (page count), `&D` (date), `&"font,style"`, `&<n>` (point size); section index 0=left, 1=center, 2=right.

```python
import aspose.cells as gc

wb = gc.Workbook("book.xlsx")
ps = wb.worksheets[0].page_setup
ps.set_header(1, "&\"Arial,Bold\"&14 Report")  # center, bold 14pt
ps.set_footer(2, "Page &P of &N")                # right: page x of y
ps.print_title_rows = "$1:$1"                    # repeat row 1 on every page
ps.print_gridlines = True                        # print cell gridlines
wb.save("out.pdf", gc.SaveFormat.PDF)
```

`print_title_columns`, `print_headings`, and the margin properties round out print behavior.

Long-running conversions (huge books, cancellation, timeouts) -> performance-and-large-files.md.

## Related
- cross-platform-deployment.md - SkiaSharp, Linux/Docker fonts, environment setup for rendering.
- performance-and-large-files.md - InterruptMonitor timeouts for long conversions.
- formulas-and-calculation.md - calculate_formula before saving computed values.
- styles-and-formatting.md - number formats that control rendered text.
- setup-and-licensing.md - evaluation watermark appearing in output.
- smart-markers-reporting.md - fill a template before converting it.
