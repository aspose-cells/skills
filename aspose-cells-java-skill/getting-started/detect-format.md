---
name: detect-format
description: Detect the format of an unknown Excel-like file (XLSX/XLS/CSV/ODS/Numbers/SpreadsheetML/encrypted) using Aspose.Cells FileFormatUtil without fully loading the workbook. Use when the user uploads a file of unknown type or extension.
applies_to: Java
keywords: [detect, format, identification, sniff, magic-bytes, encrypted]
---

# Detect file format

`FileFormatUtil.detectFileFormat(...)` reads the file's signature, then
returns a `FileFormatInfo` containing the format, encryption status,
and (for encrypted OOXML) the detected encryption standard — without
constructing a `Workbook`.

## Imports

```java
import com.aspose.cells.*;
```

## Quick example — by path

```java
FileFormatInfo info = FileFormatUtil.detectFileFormat("upload.bin");

System.out.println("Format:    " + info.getLoadFormat());  // e.g. XLSX
System.out.println("Encrypted: " + info.isEncrypted());

if (info.isEncrypted() && info.getFileFormatType() == FileFormatType.OOXML) {
    LoadOptions opt = new LoadOptions();
    opt.setPassword("guess-or-prompted");
    Workbook wb = new Workbook(new FileInputStream("upload.bin"), opt);
}
```

## Quick example — by stream

```java
try (InputStream in = new FileInputStream("upload.bin")) {
    FileFormatInfo info = FileFormatUtil.detectFileFormat(in);
    if (info.getLoadFormat() == LoadFormat.CSV) {
        // we know the contents are CSV
    }
}
```

## Pitfalls

- **CSV is detected by content, not extension** — works for `.csv`,
  `.txt`, or even no extension at all, but only if the content has
  the CSV signature.
- **`isEncrypted()` alone is not enough** — `FileFormatType` tells
  you if it's the modern OOXML encryption (AESCrypto) or the legacy
  binary encryption, which need different password/options paths.
- **Detection reads the file header** — keep the stream open until
  the `Workbook` is constructed if you intend to load the same bytes.
- **SpreadsheetML** — returns `LoadFormat.SPREADSHEET_ML` for the
  `.xml` Excel dialect, distinct from generic XML.

## Related

- [open-file.md](open-file.md)
- [security/encryption-decryption.md](../security/encryption-decryption.md)
