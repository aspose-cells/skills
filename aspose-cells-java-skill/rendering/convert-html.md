---
name: convert-html
description: Export an Aspose.Cells for Java workbook to HTML — HtmlSaveOptions for export area, single-file vs folder, embedded images, hidden sheet visibility, CSS class hooks, tooltip on cells. Use when the user wants HTML output of a workbook or sheet.
applies_to: Java
keywords: [HTML, HtmlSaveOptions, single file, hidden worksheet, css class, export]
---

# Convert workbook to HTML

## Imports

```java
import com.aspose.cells.*;
```

## Quick example — single HTML file

```java
Workbook wb = new Workbook("input.xlsx");
HtmlSaveOptions opts = new HtmlSaveOptions();
opts.setHtmlCrossStringType(HtmlCrossType.DEFAULT);
opts.setExportImagesAsBase64(true);          // inline images for portability

wb.save("out.html", opts);
```

By default Aspose writes one HTML file with embedded images. Use
`opts.setExportImagesAsBase64(false)` plus
`opts.setFilePathProvider(IFilePathProvider)` to split into multiple
files (or `setSaveAsSingleFile(false)` to produce one file per sheet).

## Sheet selection

```java
// Default is HTML when constructed as HtmlSaveOptions; only the active
// sheet is exported. Use setExportActiveWorksheetOnly(true) for clarity.
```

## Export only one sheet

```java
HtmlSaveOptions hs = new HtmlSaveOptions();
hs.setExportActiveWorksheetOnly(true);
wb.save("one-sheet.html", hs);
```

## Cell tooltip on hover

```java
HtmlSaveOptions htmlOpts = new HtmlSaveOptions();
htmlOpts.setAddTooltipText(true);
```

When enabled, the cell's full text becomes the tooltip — useful for
truncated cell values.

## Customise CSS class hooks

Aspose emits `<td class="x105y9">` style hooks so CSS can target cells
by row × column.

```java
HtmlSaveOptions cssOpts = new HtmlSaveOptions();
cssOpts.setTableCssId("report");
```

## Pitfalls

- **`ExportImagesAsBase64(true)` produces much larger HTML files**
  but a single self-contained file with no external assets.
- **Multi-sheet exports** — if not setting
  `setExportActiveWorksheetOnly(true)`, every visible sheet becomes
  its own page with navigation tabs.
- **Link targets** are relative to the output HTML's directory;
  configure `opts.setFilePathProvider(IFilePathProvider)` if the
  folder differs from the default.
- **`setEncoding`** must be set *before* `save(...)`. UTF-8 is the
  only safe default for multilingual content.
- **Hidden sheets** with `setExportHiddenWorksheet(true)` are
  included — only useful for data tables hidden from end users.

## Related

- [worksheet-image.md](worksheet-image.md)
- [convert-to-pdf.md](convert-to-pdf.md)
- [../_shared/fonts-linux-docker.md](../_shared/fonts-linux-docker.md)
