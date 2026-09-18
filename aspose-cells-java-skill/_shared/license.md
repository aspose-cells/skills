---
name: license
description: Apply the Aspose.Cells for Java license — classic License.setLicense or Metered.setMeteredKey. Use once per JVM at application start. Skipping this causes an evaluation watermark and a 100-file limit.
applies_to: Java
keywords: [license, licensing, metered, evaluation, watermark]
---

# License

Set the license **once** at application start, before constructing any
`Workbook` that will be saved or exported. Without a license, every
output file is stamped "Evaluation Only" and the API throws after 100
files.

## Imports

```java
import com.aspose.cells.License;
import com.aspose.cells.Metered;
```

## Classic license (most customers)

Place `Aspose.Cells.Java.lic` next to the JAR / in the working
directory:

```java
License lic = new License();
lic.setLicense("Aspose.Cells.Java.lic");

// Verify before saving to be sure:
Workbook wb = new Workbook();
if (!wb.isLicensed()) {
    throw new IllegalStateException("License not applied");
}
```

From a stream (classpath, DB, secrets manager):

```java
try (FileInputStream fis = new FileInputStream("/path/Aspose.Cells.Java.lic")) {
    new License().setLicense(fis);
}
```

## Metered (pay-as-you-go)

Use only when the customer has a metered subscription — never mix
classic and metered in the same JVM.

```java
Metered metered = new Metered();
metered.setMeteredKey("public-key", "private-key");

Workbook wb = new Workbook();
System.out.println(wb.isLicensed());
```

## Pitfalls

- **License file corruption** — even one stray newline invalidates
  the XML; no exception is thrown. Re-download from the customer
  portal; treat the file as opaque.
- **Web apps** — call `setLicense` in `ServletContextListener` (or
  equivalent) so it runs before any request handler creates a
  `Workbook`.
- **Multiple Aspose products** — each product needs its own
  `License`. Calling `License.setLicense` only covers `com.aspose.cells`
  — for Words/PDF/Email you need separate license files and classes.
- **Maven shade / Spring Boot fat jar** — license must be on the
  classpath root. Default path lookup checks (1) explicit file, (2)
  JAR folder, (3) classpath root.

## Related

- [getting-started/create-workbook.md](../getting-started/create-workbook.md)
- [getting-started/maven-gradle-dependencies.md](../getting-started/maven-gradle-dependencies.md)
