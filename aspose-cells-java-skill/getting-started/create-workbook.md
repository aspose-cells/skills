---
name: create-workbook
description: Create a new empty or template-based Aspose.Cells for Java Workbook in memory. Use when the user wants to start a workbook from scratch — code generates a fresh XLSX, not loading from disk.
applies_to: Java
keywords: [workbook, new, create, empty, template]
---

# Create a new workbook

The default constructor `new Workbook()` produces one worksheet
(`"Sheet1"`); the format-fit constructor picks `XLSX`.

## Imports

```java
import com.aspose.cells.*;
```

## Quick example — empty XLSX

```java
Workbook workbook = new Workbook();           // XLSX by default
Worksheet sheet  = workbook.getWorksheets().get(0);
sheet.getCells().get("A1").setValue("Hello");
workbook.save("out.xlsx");
```

## Quick example — old XLS (97-2003)

```java
Workbook workbook = new Workbook();
workbook.save("legacy.xls", SaveFormat.EXCEL_97_TO_2003);
```

## Quick example — ODS / CSV / PDF

```java
Workbook workbook = new Workbook();
workbook.save("sheet.ods", SaveFormat.ODS);
workbook.save("sheet.fods", SaveFormat.FODS);
```

## Template / seed data → XLSX

If you have a pre-prepared XLSX on the classpath that contains named
ranges, validated cells, or styles:

```java
try (InputStream template = Thread.currentThread().getContextClassLoader()
                                   .getResourceAsStream("/report-template.xlsx")) {
    Workbook workbook = new Workbook(template);
    // populate cells, then:
    workbook.save("report.xlsx");
}
```

## Pitfalls

- **`new Workbook(File)` / `(String)` is for loading**, not creating.
  Use `new Workbook()` for an empty workbook; pass a `FileInputStream`
  only when loading an existing template.
- **First sheet exists** — `getWorksheets().get(0)` is never null.
  `add()` creates additional sheets, never the first one.
- **PDF as the initial format is not supported** — create as XLSX,
  then save as PDF through `rendering/convert-to-pdf.md`.
- **Apply license first** (see `../_shared/license.md`) — otherwise
  each save tacks on the evaluation watermark.

## Related

- [open-file.md](open-file.md)
- [save-export.md](save-export.md)
- [_shared/license.md](../_shared/license.md)
