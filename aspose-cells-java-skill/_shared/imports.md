---
name: imports
description: Standard Java imports for Aspose.Cells for Java. Use when generating any code block — paste this list verbatim unless the task strictly needs fewer imports.
applies_to: Java
keywords: [imports, package]
---

# Imports — single source

```java
import com.aspose.cells.*;
```

That is the **only** import needed for ~95 % of Aspose.Cells tasks.
The package already exposes every public type — `Workbook`,
`Worksheet`, `Cell`, `PdfSaveOptions`, `Chart`, `PivotTable`,
`TxtSaveOptions`, etc.

Additional imports (rare):

| Task                              | Additional import                                           |
| --------------------------------- | ----------------------------------------------------------- |
| File I/O helpers                  | `import java.io.FileInputStream;`                           |
| Streams / byte arrays             | `import java.io.ByteArrayOutputStream;` etc.                |

## Pitfalls

- **Never import BouncyCastle types.** Aspose.Cells calls BouncyCastle
  internally and registers the provider itself — the jars only need to
  be on the classpath.
- **Do not combine** `com.aspose.cells.*` with `org.apache.poi.*` in
  the same file — both define a `Workbook` class.
- **Do not import** `com.aspose.gridweb.*` unless the user is
  building a GridWeb control — the `License` class there is
  different from `com.aspose.cells.License`.
- The properties class for some optional features (cells data,
  metadata) lives in `com.aspose.cells.metadata.*` or
  `com.aspose.cells.system.*` — import those only when their types
  actually appear in your code.
