---
name: convert-to-pdf
description: Convert an Aspose.Cells for Java workbook to PDF — PdfSaveOptions for compliance, page layout, fonts, bookmarks, security, watermark, attachments, one-page-per-sheet. Use when the user wants any PDF output of a workbook.
applies_to: Java
keywords: [PDF, PdfSaveOptions, PDFA1A, one page per sheet, bookmark, font, watermark, PDFA]
---

# Convert workbook to PDF

## Imports

```java
import com.aspose.cells.*;
```

## Quick example

```java
Workbook wb = new Workbook("input.xlsx");
PdfSaveOptions opts = new PdfSaveOptions();
opts.setCompliance(PdfCompliance.PDF_A_1_A);    // optional: PDF/A compliant
opts.setOnePagePerSheet(true);                   // one sheet per page

wb.save("out.pdf", opts);
```

## Selected options (PdfSaveOptions)

| Method                                  | Use                                            |
| --------------------------------------- | ---------------------------------------------- |
| `setCompliance(...)`                     | `PDF_A_1_A`, `PDF_A_1_B`, `PDF_A_2_A`, `PDF_15` etc. |
| `setOnePagePerSheet(boolean)`           | Force one sheet per page                       |
| `setAllColumnsInOnePagePerSheet(true)`  | All columns on one page width                  |
| `setPageIndex(int)`                     | Start saving from page N                       |
| `setPageCount(int)`                     | Limit number of pages                          |
| `setPageSavingCallback(IPageSavingCallback)` | Page-save callbacks                            |
| `setEmbedStandardWindowsFonts(true)`    | Embed standard Windows fonts (subset, used for PDF/A) |
| `setWarningCallback(IWarningCallback)`  | Surface missing-font warnings (warn/fail via the IWarningCallback) |
| `setBookmark(PdfBookmarkEntry)`         | Add a single PDF bookmark entry (repeat per entry) |
| `setSecurityOptions(...)`               | Owner / user password on resulting PDF         |

```java
Workbook wb = new Workbook("input.xlsx");
PdfSaveOptions opts = new PdfSaveOptions();
opts.setOnePagePerSheet(true);
opts.setAllColumnsInOnePagePerSheet(true);     // long tables fit horizontally
opts.setDefaultFont("Arial");                   // font used when a sheet font is missing
wb.save("out.pdf", opts);
```

## Bookmarks (PDF outline)

```java
Workbook wb = new Workbook("input.xlsx");
Worksheet ws = wb.getWorksheets().get(0);
PdfBookmarkEntry root = new PdfBookmarkEntry();
root.setText("Quarterly");
root.setDestination(ws.getCells().get("A1"));

PdfBookmarkEntry q1 = new PdfBookmarkEntry();
q1.setText("Q1");
q1.setDestination(ws.getCells().get("A2"));

ArrayList<PdfBookmarkEntry> subs = new ArrayList<>();
subs.add(q1);
root.setSubEntry(subs);                          // setSubEntry accepts a typed list
PdfSaveOptions opts = new PdfSaveOptions();
opts.setBookmark(root);
wb.save("out.pdf", opts);
```

## Password-protect the PDF

```java
Workbook wb = new Workbook("input.xlsx");
PdfSaveOptions opts = new PdfSaveOptions();
PdfSecurityOptions sec = new PdfSecurityOptions();
sec.setUserPassword("view-pwd");
sec.setOwnerPassword("print-pwd");
sec.setPrintPermission(true);
sec.setExtractContentPermission(true);          // allow copy / paste

opts.setSecurityOptions(sec);
wb.save("out.pdf", opts);
```

## Add a watermark

`PdfSaveOptions.setWatermark(RenderingWatermark)` stamps every page —
no need to insert a picture into the sheet.

```java
Workbook wb = new Workbook("input.xlsx");

RenderingFont font = new RenderingFont("Calibri", 68);
font.setItalic(true);
font.setBold(true);
font.setColor(com.aspose.cells.Color.getBlue());

RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font);
watermark.setHAlignment(TextAlignmentType.CENTER);
watermark.setVAlignment(TextAlignmentType.CENTER);
watermark.setRotation(30);
watermark.setOpacity(0.6f);
watermark.setScaleToPagePercent(50);          // percent of page size

PdfSaveOptions opts = new PdfSaveOptions();
opts.setWatermark(watermark);
wb.save("out.pdf", opts);
```

For an image watermark, build it from the image bytes instead — the
alignment, rotation, opacity and scale setters are the same:

```java
byte[] imageBytes = java.nio.file.Files.readAllBytes(
        java.nio.file.Paths.get("/path/to/watermark.png"));

RenderingWatermark watermark = new RenderingWatermark(imageBytes);
watermark.setBackground(true);                // behind the content
watermark.setOffsetX(100);
watermark.setOffsetY(200);
watermark.setOpacity(0.6f);

PdfSaveOptions opts = new PdfSaveOptions();
opts.setWatermark(watermark);
```

## Pitfalls

- **Always pass `PdfSaveOptions`**. `wb.save("x.pdf")` ignores
  options and uses defaults.
- **Linux Docker** — install fonts before exporting PDF; missing
  fonts produce "tofu" squares in the PDF. See
  `../_shared/fonts-linux-docker.md`.
- **PDF/A-1A strict conformance** — `setEmbedStandardWindowsFonts(true)`
  otherwise the validator will reject.
- **One-page-per-sheet can dramatically widen output** — combine
  with `setAllColumnsInOnePagePerSheet(true)` for long tables.
- **Page index bounds** — `setPageCount` cannot exceed the actual
  rendered pages; setting an unreachable count is silently clamped.

## Related

- [convert-html.md](convert-html.md)
- [worksheet-image.md](worksheet-image.md)
- [../_shared/fonts-linux-docker.md](../_shared/fonts-linux-docker.md)
- [../getting-started/save-export.md](../getting-started/save-export.md)
