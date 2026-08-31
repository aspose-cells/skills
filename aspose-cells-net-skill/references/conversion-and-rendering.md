# Conversion, Rendering, and Saving

Read when saving to any format, converting Excel to PDF/HTML/image/JSON/CSV, controlling PDF page layout, a downloaded file arrives corrupt, or choosing xls vs xlsx.

Conversion means load a workbook in one format and `Save` it in another. Rendering (PDF/image/XPS/HTML) is layout-sensitive: output follows each worksheet's `PageSetup` unless you override it. PDF, XPS, and image formats are save-only render targets - they cannot be loaded back into a `Workbook`.

## Pitfalls

### Excel-to-PDF splits pages wrong or content overflows
- Symptom: columns spill onto extra pages, a wide table breaks mid-row, or one sheet becomes dozens of PDF pages.
- Cause: paging defaults to each sheet's `PageSetup` (paper size, margins, manual breaks). Fit-to-page is opt-in, not automatic.
- Fix: force layout via `PdfSaveOptions` or per-sheet `PageSetup`:
  - `PdfSaveOptions.OnePagePerSheet = true` - the entire sheet on one page.
  - `PdfSaveOptions.AllColumnsInOnePagePerSheet = true` - all columns one page wide, rows may overflow.
  - `sheet.PageSetup.FitToPagesWide = 1; FitToPagesTall = 0` - scale to 1 page wide, unlimited tall.
  - Constrain output with `PageSetup.PrintArea`, `Orientation`, `PaperSize`. See the PDF how-to below.

### Missing fonts substitute silently and shift PDF/image layout
- Symptom: rendered text differs from Excel - wrong glyphs, changed widths, misaligned columns. No exception is thrown.
- Cause: the render font must exist on the machine. If absent, Aspose.Cells substitutes the closest available font (finally Arial), which changes metrics.
- Fix: register the fonts and set a deterministic fallback:
  - `PdfSaveOptions.DefaultFont` / `ImageOrPrintOptions.DefaultFont` - fallback face for unmapped glyphs.
  - `PdfSaveOptions.IsFontSubstitutionCharGranularity = true` - substitute per character, not per run.
  - `FontConfigs.SetFontFolder(dir, true)` at startup, before any render. Cross-platform/Docker font setup -> cross-platform-deployment.md.
  - Detect substitutions with an `IWarningCallback` (below) instead of eyeballing output.

### Downloaded xlsx/pdf arrives corrupt or zero-length (web apps)
- Symptom: the browser downloads a file that will not open; bytes are truncated or empty.
- Cause: after `Save(stream, ...)` the stream `Position` sits at the end, so code that copies "from the current position" sends nothing; or the declared `SaveFormat`, MIME type, and extension disagree.
- Fix: reset `ms.Position = 0` before reading (or use `ms.ToArray()`), and keep format/MIME/extension consistent. See the web-download how-to.

### SaveFormat.Excel97To2003 (xls) truncates or fails on large data
- Symptom: rows past 65,536 or columns past 256 vanish, or Save fails, when writing xls.
- Cause: the legacy .xls (BIFF8) grid is capped at 65,536 rows x 256 columns; xlsx/xlsb allow 1,048,576 x 16,384.
- Fix: save large data as `SaveFormat.Xlsx` (or `Xlsb` for size and speed). Use `Excel97To2003` only when a consumer truly requires .xls, and keep data within the legacy limits.

### PDF hairlines look thicker in Adobe Reader
- Symptom: borders/gridlines render heavier in Adobe Reader than in Excel or other viewers.
- Cause: Adobe Reader's "Enhance thin lines" / "Smooth line art" display settings, not an output defect. The PDF geometry is correct.
- Fix: nothing in code. In Adobe Reader: Edit > Preferences > Page Display, uncheck "Smooth line art" and "Enhance thin lines". Other viewers are unaffected.

## Save formats and Save overloads

`Workbook.Save` picks the target from the `SaveFormat` enum (or a `*SaveOptions` object).

| Target | SaveFormat member | Notes |
| :- | :- | :- |
| xlsx | `Xlsx` | default modern workbook |
| xlsm | `Xlsm` | macro-enabled |
| xlsb | `Xlsb` | binary; smallest and fastest for big data |
| xls | `Excel97To2003` | legacy; 65,536 x 256 limit |
| csv / tsv | `Csv` / `Tsv` | active sheet only by default |
| ods | `Ods` | OpenDocument Calc |
| pdf | `Pdf` | pair with `PdfSaveOptions` |
| html / mht | `Html` / `MHtml` | pair with `HtmlSaveOptions` |
| json | `Json` | pair with `JsonSaveOptions` |

`Save(path)` infers the format from the file extension; pass a `SaveFormat` explicitly when the extension is missing or ambiguous.

```csharp
var wb = new Workbook("book.xlsx");
wb.Save("out.pdf");                        // format inferred from ".pdf"
wb.Save("out.data", SaveFormat.Xlsx);      // explicit format
using var fs = File.Create("copy.xlsx");
wb.Save(fs, SaveFormat.Xlsx);              // to a stream
```

Also `Save(path, SaveOptions)` and `Save(Stream, SaveOptions)` for fine-grained control.

## Excel to PDF

Set options on `PdfSaveOptions`. If the sheet has formulas, call `CalculateFormula()` (or set `PdfSaveOptions.CalculateFormula = true`) before saving so rendered values are current.

```csharp
var wb = new Workbook("book.xlsx");
wb.CalculateFormula();                       // refresh formula results first
var opts = new PdfSaveOptions();
opts.OnePagePerSheet = true;                 // whole sheet -> one page
opts.Compliance = PdfCompliance.PdfA1b;      // archival PDF/A-1b (else PDF 1.7)
wb.Save("out.pdf", opts);
```

Control page layout per sheet through `PageSetup`:

```csharp
var wb = new Workbook("book.xlsx");
var ps = wb.Worksheets[0].PageSetup;
ps.FitToPagesWide = 1;                        // all columns across 1 page
ps.FitToPagesTall = 0;                        // 0 = as many row-pages as needed
ps.PrintArea = "A1:H200";
ps.Orientation = PageOrientationType.Landscape;
ps.PaperSize = PaperSizeType.PaperA4;
wb.Save("out.pdf", SaveFormat.Pdf);
```

Render only chosen sheets with `PdfSaveOptions.SheetSet`:

```csharp
var wb = new Workbook("book.xlsx");
var opts = new PdfSaveOptions();
opts.SheetSet = SheetSet.Active;              // or SheetSet.All (includes hidden)
// opts.SheetSet = new SheetSet(new[] { 0, 2 });  // sheets 0 and 2, in this order
wb.Save("selected.pdf", opts);
```

`PdfSaveOptions` also exposes `SecurityOptions` (encryption/permissions), `Producer`, and `CreatedTime`.

### Detect font substitution while rendering

Implement `IWarningCallback` and assign it to the save options' `WarningCallback`:

```csharp
class FontWarner : IWarningCallback
{
    public void Warning(WarningInfo info)
    {
        if (info.WarningType == WarningType.FontSubstitution)
            Console.WriteLine(info.Description);   // logs the missing font + cell
    }
    static void Run()
    {
        var wb = new Workbook("book.xlsx");
        wb.Save("out.pdf", new PdfSaveOptions { WarningCallback = new FontWarner() });
    }
}
```

## Excel to image (PNG/JPEG/TIFF)

Use `SheetRender` (one worksheet) or `WorkbookRender` (whole book), configured with `ImageOrPrintOptions`. `ImageType` defaults to PNG.

```csharp
var wb = new Workbook("book.xlsx");
var opts = new ImageOrPrintOptions();
opts.ImageType = ImageType.Png;
opts.HorizontalResolution = 200;              // DPI; raise for print quality
opts.VerticalResolution = 200;
opts.OnePagePerSheet = true;
var sr = new SheetRender(wb.Worksheets[0], opts);
for (int p = 0; p < sr.PageCount; p++)
    sr.ToImage(p, $"page{p}.png");            // one file per page
```

Whole workbook to a single multi-page TIFF:

```csharp
var wb = new Workbook("book.xlsx");
var opts = new ImageOrPrintOptions();
opts.ImageType = ImageType.Tiff;
var wr = new WorkbookRender(wb, opts);
wr.ToImage("workbook.tiff");                  // all pages in one TIFF
```

`SheetRender.ToTiff(path)` writes a multi-page TIFF for a single sheet. `ToImage` overloads also accept a `Stream`.

## Excel to HTML

```csharp
var wb = new Workbook("book.xlsx");
var opts = new HtmlSaveOptions();
opts.ExportImagesAsBase64 = true;             // self-contained single .html
opts.ExportActiveWorksheetOnly = true;        // omit to export every sheet
wb.Save("out.html", opts);
```

`HtmlSaveOptions.ImageOptions` (an `ImageOrPrintOptions`) tunes how shapes/charts rasterize; `ExportGridLines` and `PageTitle` are also available.

## Excel to and from JSON

Export the active sheet with `SaveFormat.Json`, or a range with `JsonUtility`:

```csharp
var wb = new Workbook("book.xlsx");
wb.Save("out.json", SaveFormat.Json);         // active sheet
Aspose.Cells.Range range = wb.Worksheets[0].Cells.CreateRange("A1:C10");
string json = JsonUtility.ExportRangeToJson(range, new JsonSaveOptions());
```

Import a JSON string into cells with a layout, or load a whole `.json` file:

```csharp
var wb = new Workbook();
var cells = wb.Worksheets[0].Cells;
string data = File.ReadAllText("data.json");
JsonUtility.ImportData(data, cells, 0, 0, new JsonLayoutOptions());
var whole = new Workbook("data.json", new LoadOptions(LoadFormat.Json));
```

## Excel to CSV/TSV/text

CSV/TSV/TXT save the active sheet only unless told otherwise. Use `TxtSaveOptions` for delimiter and multi-sheet control.

```csharp
var wb = new Workbook("book.xlsx");
wb.Save("out.csv", SaveFormat.Csv);           // active sheet, comma-separated
var txt = new TxtSaveOptions(SaveFormat.Csv);
txt.Separator = ';';                          // custom delimiter
txt.ExportAllSheets = true;                    // concatenate all sheets
wb.Save("all.csv", txt);
```

## LowCode one-call converters

The `Aspose.Cells.LowCode` namespace runs whole-file conversions without building a `Workbook`. Fully qualify these types (they are not in the default usings). Each exposes a static `Process`.

```csharp
Aspose.Cells.LowCode.PdfConverter.Process("book.xlsx", "out.pdf");
Aspose.Cells.LowCode.SpreadsheetConverter.Process("book.xlsx", "out.xlsb");
Aspose.Cells.LowCode.SpreadsheetLocker.Process("book.xlsx", "locked.xlsx", "pw", "pw");
```

For finer control pass `LowCodeLoadOptions` + `LowCodeSaveOptions`, e.g. `PdfConverter.Process(load, save)`. Siblings: `HtmlConverter`, `ImageConverter`, `JsonConverter`, `TextConverter`, `SpreadsheetMerger.Process(string[] inputs, string output)`, and `SpreadsheetSplitter.Process(input, output)`.

## Web download (framework-neutral)

Save to a `MemoryStream`, rewind, then hand the bytes to the response. MIME types: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` for xlsx, `application/vnd.ms-excel` for xls, `application/pdf` for pdf.

```csharp
var wb = new Workbook("book.xlsx");
var ms = new MemoryStream();
wb.Save(ms, SaveFormat.Xlsx);                 // explicit format; match MIME + name
ms.Position = 0;                              // Save leaves Position at the end
byte[] bytes = ms.ToArray();                  // or copy the rewound stream out
```

In ASP.NET Core, return `File(ms, mime, "report.xlsx")` after `ms.Position = 0` (or return `bytes`). The legacy `Workbook.Save(HttpResponse, name, ContentDisposition, options)` overload exists only on .NET Framework builds, not .NET 5+.

## Page setup essentials

Header/footer scripts use `&P` (page), `&N` (page count), `&D` (date), `&"font,style"`, `&<n>` (point size); section index 0=left, 1=center, 2=right.

```csharp
var wb = new Workbook("book.xlsx");
var ps = wb.Worksheets[0].PageSetup;
ps.SetHeader(1, "&\"Arial,Bold\"&14 Report"); // center, bold 14pt
ps.SetFooter(2, "Page &P of &N");             // right: page x of y
ps.PrintTitleRows = "$1:$1";                    // repeat row 1 on every page
ps.PrintGridlines = true;                       // print cell gridlines
wb.Save("out.pdf", SaveFormat.Pdf);
```

`PrintTitleColumns`, `PrintHeadings`, and the margin properties round out print behavior.

Long-running conversions (huge books, cancellation, timeouts) -> performance-and-large-files.md.

## Related
- cross-platform-deployment.md - SkiaSharp, Linux/Docker fonts, environment setup for rendering.
- performance-and-large-files.md - InterruptMonitor timeouts for long conversions.
- formulas-and-calculation.md - CalculateFormula before saving computed values.
- styles-and-formatting.md - number formats that control rendered text.
- setup-and-licensing.md - evaluation watermark appearing in output.
- smart-markers-reporting.md - fill a template before converting it.
