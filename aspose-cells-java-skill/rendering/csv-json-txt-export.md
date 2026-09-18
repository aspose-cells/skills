---
name: csv-json-txt-export
description: Export an Aspose.Cells for Java workbook to CSV, TSV, TXT, or JSON — TxtSaveOptions for delimiter, encoding, export-all-sheets; JsonSaveOptions for nested-vs-flat, nesting structure. Use when the user wants plain-text or JSON output.
applies_to: Java
keywords: [CSV, TSV, TXT, JSON, TxtSaveOptions, JsonSaveOptions, export, export all sheets]
---

# CSV / TSV / TXT / JSON export

## Imports

```java
import com.aspose.cells.*;
```

## CSV (active sheet) — quick example

```java
Workbook wb = new Workbook("input.xlsx");
TxtSaveOptions opts = new TxtSaveOptions();
opts.setSeparator(',');
wb.save("out.csv", opts);
```

## Export every sheet to one CSV each

```java
TxtSaveOptions opts = new TxtSaveOptions();
opts.setExportAllSheets(true);
wb.save("all-sheets.zip", opts);              // zip per sheet
```

`setExportAllSheets(true)` is the only way to get more than the
**active** sheet in one go; otherwise call `setActiveSheetIndex` and
save repeatedly.

## TSV / tab-separated

```java
TxtSaveOptions opts = new TxtSaveOptions();
opts.setSeparator('\t');
wb.save("out.tsv", opts);
```

## TXT (Microsoft-style flat text)

```java
Workbook wb = new Workbook();
TxtSaveOptions t = new TxtSaveOptions();
t.setSeparator('\t');                    // tab-separated
wb.save("out.txt", t);
```

## JSON export

```java
JsonSaveOptions j = new JsonSaveOptions();
j.setExportNestedStructure(true);            // nested objects for table-like data
j.setExportAsString(true);                   // numbers stay as JSON strings
wb.save("out.json", j);
```

## Pitfalls

- **Only the active sheet exports** unless
  `setExportAllSheets(true)`. The most common support ticket is
  "I exported three sheets but only got one CSV".
- **Decimal separator** on continental locales — by default the CSV
  uses `.`, not `,`, for decimals. Configure via styles if the user
  expects `,`.
- **`setEncoding`** must be set before `save()`. UTF-8 with BOM is
  needed for Excel-on-Windows to auto-detect non-ASCII.
- **Excel's CSV** can emit leading apostrophes (for `'00123`); set
  the QuotePrefix flag on the cell style to control this.
- **JSON export** produces a flat array if all columns are simple;
  use `setExportNestedStructure(true)` for nested objects based on
  repeated keys.

## Related

- [convert-to-pdf.md](convert-to-pdf.md)
- [convert-html.md](convert-html.md)
- [../getting-started/save-export.md](../getting-started/save-export.md)
