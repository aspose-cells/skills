---
name: open-file
description: Open / load an existing XLSX, XLS, CSV, ODS, SpreadsheetML, or encrypted workbook in Aspose.Cells for Java. Use when the user has a path or input stream and wants to read or modify the file.
applies_to: Java
keywords: [open, load, read, xlsx, xls, csv, ods, encrypted, loadoptions]
---

# Open an existing file

## Imports

```java
import com.aspose.cells.*;
```

## Three ways to open

```java
// 1) Path (simplest)
Workbook wb = new Workbook("/data/sales.xlsx");

// 2) Stream — recommended for uploads, classpath, S3
try (InputStream in = new FileInputStream("/data/sales.xlsx")) {
    wb = new Workbook(in);
}
```

## Force format & control parsing

When the extension does not match the bytes (CSV without `.csv`, XLSX
from another app):

```java
TxtLoadOptions opt = new TxtLoadOptions();
opt.setPassword("secret");                         // for encrypted files
opt.setMemorySetting(MemorySetting.MEMORY_PREFERENCE);  // large files
Workbook wb = new Workbook("odd-data", opt);       // pass name + opts
```

## Catch parsing warnings

```java
LoadOptions opt = new LoadOptions();
opt.setWarningCallback(new IWarningCallback() {
    @Override public void warning(WarningInfo info) {
        System.out.println(info.getDescription());
    }
});
Workbook wb = new Workbook("dirty.xlsx", opt);
```

## Read Apple's Numbers format

```java
NumbersLoadOptions opt = new NumbersLoadOptions();
Workbook wb = new Workbook("from-mac.numbers", opt);
```

## Pitfalls

- **Don't call `.close()` on the stream** before `Workbook`
  finishes — the constructor will read past the EOF and throw.
- **`LoadFormat.AUTO` is the default** but mis-classifies CSVs
  lacking `.csv` extension. Use `TxtLoadOptions` instead of
  `LoadOptions` when in doubt.
- **Encrypted XLSX** — must supply the password via `LoadOptions`,
  not by trying to open twice.
- **Stream overload is more permissive** — strings depend on JVM
  charset; streams let you specify the encoding explicitly.
- **Use `MemorySetting.MEMORY_PREFERENCE`** for files >100k rows;
  see `../_shared/large-files-memory.md`.

## Related

- [create-workbook.md](create-workbook.md)
- [save-export.md](save-export.md)
- [detect-format.md](detect-format.md)
- [../_shared/large-files-memory.md](../_shared/large-files-memory.md)
