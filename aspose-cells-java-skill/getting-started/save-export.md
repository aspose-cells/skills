---
name: save-export
description: Save or export an Aspose.Cells for Java Workbook to XLSX, XLS, PDF, HTML, CSV, JSON, ODS, image, or XPS. Use whenever the user wants to persist or convert a workbook, or asks which format to pick. Always use SaveOptions, never the overload that omits them.
applies_to: Java
keywords: [save, export, format, conversion, xlsx, xls, pdf, html, csv, json, ods]
---

# Save & export a workbook

`workbook.save(...)` accepts a path and an optional `SaveOptions` /
`SaveFormat`. **Always use the `SaveOptions` overload** for PDF, HTML,
image, and CSV — bare `save(path)` ignores per-format configuration.

## Imports

```java
import com.aspose.cells.*;
```

## Quick example — XLSX (default)

```java
wb.save("out.xlsx");
```

## Pick the format explicitly

```java
wb.save("out.xls", SaveFormat.EXCEL_97_TO_2003);
wb.save("out.ods", SaveFormat.ODS);
```

## Format-specific options (use these)

| To     | Options class                          | Skill                                       |
| ------ | -------------------------------------- | ------------------------------------------- |
| PDF    | `PdfSaveOptions`                       | `../rendering/convert-to-pdf.md`            |
| HTML   | `HtmlSaveOptions`                      | `../rendering/convert-html.md`              |
| Image  | `ImageOrPrintOptions` + `*Render`      | `../rendering/worksheet-image.md`           |
| CSV/TSV| `TxtSaveOptions`                       | `../rendering/csv-json-txt-export.md`       |
| JSON   | `JsonSaveOptions`                      | `../rendering/csv-json-txt-export.md`       |
| XLSX   | `OoxmlSaveOptions` (all OOXML formats) | this file                                   |
| XLS    | `XlsSaveOptions`                       | this file                                   |

```java
PdfSaveOptions pdf = new PdfSaveOptions();
pdf.setCompliance(PdfCompliance.PDF_A_1_A);
pdf.setOnePagePerSheet(true);
wb.save("out.pdf", pdf);
```

## Encrypt while saving XLSX

Three workbook-security mechanisms are independent — file
**encryption** (this snippet), **write protection** (open
read-only), and **structure protection** (block add/rename/delete
sheets). See `security/encryption-decryption.md` for the full
picture.

```java
// Encrypt the file's bytes — opening requires this password.
wb.getSettings().setPassword("open-pwd");
wb.save("encrypted.xlsx");
```

## Track conversion progress

```java
PdfSaveOptions pdf = new PdfSaveOptions();
pdf.setPageSavingCallback(new IPageSavingCallback() {
    @Override public void pageStartSaving(PageStartSavingArgs args) {
        System.out.println("saving page " + args.getPageIndex());
    }
    @Override public void pageEndSaving(PageEndSavingArgs args) { }
});
wb.save("out.pdf", pdf);
```

## Pitfalls

- **`workbook.save("x.pdf")` ignores PDF options** — always pass
  `PdfSaveOptions` when you need compliance, compression, or page
  grouping.
- **Output file extension must match the format** — Aspose inspects
  it. Mismatch produces "this file format is not supported".
- **CSV export by default writes only the active sheet** — see
  `../rendering/csv-json-txt-export.md`.
- **`save()` writes through to disk; close is not needed.** For
  in-memory pipelines use the `wb.save(OutputStream, saveOptions)`
  overload with a `ByteArrayOutputStream`.
- **Always `calculateFormula()` before saving** if the user edits
  formulas — see `../formulas/calculate-formulas.md`.

## Related

- [rendering/convert-to-pdf.md](../rendering/convert-to-pdf.md)
- [rendering/convert-html.md](../rendering/convert-html.md)
- [rendering/csv-json-txt-export.md](../rendering/csv-json-txt-export.md)
- [_shared/license.md](../_shared/license.md)
