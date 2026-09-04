---
name: setup-and-licensing
description: Install Aspose.Cells for Java, apply and verify a license (classic .lic file or Metered), and avoid the silent-evaluation-mode traps (subscription expiry, missing file path, weekly-trial bundle). Use when output shows an evaluation watermark, a 100-file limit, or an "Evaluation Only" sheet, or when wiring a license for a new server.
applies_to: Java
keywords: [license, setLicense, License, Metered, isLicensed, subscriptionExpireDate, watermark, evaluation]
---

# Setup and licensing

Read this when installing Aspose.Cells, applying or verifying a
license, or when output shows an evaluation watermark ("Evaluation
Only. Created with Aspose.Cells") or a truncated file. The license
is a digitally signed XML file that you must set **once per JVM** at
application start, before constructing any `Workbook`.

See [_shared/license.md](../_shared/license.md) for the short
recipe. This file covers the traps that the short recipe doesn't.

## Imports

```java
import com.aspose.cells.License;
import com.aspose.cells.Metered;
import com.aspose.cells.Workbook;
```

## Install (Maven)

License file is separate from the JAR — install via Maven, then drop
the `.lic` next to the JAR or on the classpath.

```xml
<dependency>
    <groupId>com.aspose</groupId>
    <artifactId>aspose-cells</artifactId>
    <version>26.8</version>
</dependency>
```

The license comes from the purchase portal or the free 30-day
temporary license at <https://purchase.aspose.com>. Microsoft Excel
is NOT required on the machine.

## Pitfalls

### Evaluation watermark though you own a license

Symptom: an extra worksheet titled "Evaluation Only. Created with
Aspose.Cells" is injected, PDF/CSV files are truncated or stamped
with a watermark, and opening more than 100 files throws. `Workbook.
isLicensed()` returns `false`. Four distinct causes:

a. **SetLicense not called, called too late, or pointed at the wrong
   class.** Apply the license once at JVM start, before any
   `Workbook` is constructed. Use `com.aspose.cells.License` — *not*
   `com.aspose.words.License` or any other Aspose product's License
   class. Each product has its own `.lic` file.

```java
// Call ONCE at startup, BEFORE constructing any Workbook.
License lic = new License();
lic.setLicense("Aspose.Cells.Java.lic");
Workbook wb = new Workbook();                    // licensed from this point
wb.save("licensed.xlsx");
```

b. **License file not found at runtime** (works on the dev box,
   watermark on the server / inside a container). See "License file
   not found at runtime" below.

c. **Subscription-expiry trap** ("worked yesterday, watermark after
   a package upgrade"). Your `.lic` carries an expiry date. The
   file unlocks every Aspose.Cells JAR released **on or before** that
   date — but a JAR released **after** the date silently reverts to
   evaluation. Upgrading the Maven dependency to a newer version
   mid-subscription is the classic trigger. Fix: pin the
   `aspose-cells` version to one released within your subscription
   window, or renew the subscription. Check the expiry at runtime:

```java
import com.aspose.cells.License;
import java.util.Date;

Date expiry = License.getSubscriptionExpireDate();
if (expiry != null && expiry.before(new Date())) {
    throw new IllegalStateException(
        "Aspose.Cells license expired on " + expiry
      + " — newer JAR releases will run in evaluation mode.");
}
```

d. **Edited or re-saved `.lic` file.** The file is digitally signed.
   Any change — adding a newline, re-encoding, opening and
   re-saving in an editor — invalidates the signature. Use it
   byte-for-byte as delivered; never modify or reformat it.

### License file not found at runtime

Symptom: passing just the file name works locally but produces a
watermark after deployment, or `setLicense` throws.

Cause: the `.lic` is not on any path the JVM resolves. `setLicense(String)`
looks in this order:

1. The explicit path you pass.
2. The folder containing the `aspose-cells.jar` on the JVM classpath.
3. The root of the JVM classpath (works for `src/main/resources`).

Fix (recommended for servers / containers): embed the license as a
classpath resource and load it via `InputStream`. Then it can
never be lost or mispathed.

```java
// src/main/resources/Aspose.Cells.Java.lic
try (java.io.InputStream lic =
        Thread.currentThread().getContextClassLoader()
                .getResourceAsStream("Aspose.Cells.Java.lic")) {
    if (lic == null) {
        throw new IllegalStateException(
            "License not found on classpath: Aspose.Cells.Java.lic");
    }
    new License().setLicense(lic);
}
```

The Java classloader resolves the resource relative to the
classpath root, so do NOT prefix a leading `/`. Keep the file in
`src/main/resources` so Maven Shade / Spring Boot fat-jar builds
place it on the classpath root.

When you control the bytes (blob store, secret manager), pass the
`InputStream` directly. The point is: any `InputStream` works —
load it from wherever you store the license, then hand it to
`setLicense`.

```java
// Read from a known path. Replace with your blob-store / secret-manager
// client call (e.g. Files.newInputStream, SecretClient.getSecret(...).getInputStream()).
try (java.io.InputStream lic =
        java.nio.file.Files.newInputStream(
            java.nio.file.Paths.get("/etc/aspose/Aspose.Cells.Java.lic"))) {
    new License().setLicense(lic);
}
```

If you have to load from a real path on disk, use the `setLicense(String)`
overload with the full absolute path:

```java
new License().setLicense("/etc/aspose/Aspose.Cells.Java.lic");
```

## Verifying the license loaded

A missing license file makes `setLicense` throw a
`FileNotFoundException` ("Cannot find license"); a corrupt or
modified file throws an `Exception` ("Corrupted License file.")
— verified against 26.7. Do not swallow either. Always assert
`isLicensed()` after construction as a smoke test.

```java
License lic = new License();
try {
    lic.setLicense("Aspose.Cells.Java.lic");
} catch (Exception ex) {
    throw new IllegalStateException("License NOT applied: " + ex.getMessage(), ex);
}
Workbook wb = new Workbook();
if (!wb.isLicensed()) {
    throw new IllegalStateException("License did not apply — setLicense ran but isLicensed() is false");
}
```

## Evaluation version limitations

Recognize these as "no license loaded" symptoms, not bugs. Without
a license Aspose.Cells is fully functional except:

- An extra worksheet stamped "Evaluation Only. Created with
  Aspose.Cells" is injected and forced to be the active sheet.
- Only 100 files may be opened per process; opening more throws.
- Plain-text output (CSV, TSV) appends evaluation text and writes
  only the FIRST worksheet.
- PDF and image output get an evaluation watermark banner at the top.

All disappear once a valid license is applied before the first
`Workbook`.

## Metered (pay-as-you-go)

An alternative to a `.lic` file: bill by usage instead of a fixed
license. Call `setMeteredKey(publicKey, privateKey)` on a `Metered`
instance at startup. Requires network access — consumption is
reported to Aspose — and meters by processed file size.

```java
import com.aspose.cells.Metered;
import com.aspose.cells.Workbook;

Metered metered = new Metered();
metered.setMeteredKey("<public key>", "<private key>");

Workbook wb = new Workbook();
System.out.println("licensed via Metered: " + wb.isLicensed());
```

Never mix classic and Metered in the same JVM — pick one.

## Web apps

Call `setLicense` in a `ServletContextListener` (or equivalent) so
it runs before any request handler creates a `Workbook`. Putting
it in a `@PostConstruct` is too late if another bean has already
constructed a workbook at startup.

## Related

- [_shared/license.md](../_shared/license.md) — short recipe.
- [getting-started/maven-gradle-dependencies.md](maven-gradle-dependencies.md)
- [getting-started/create-workbook.md](create-workbook.md)
- [_shared/large-files-memory.md](../_shared/large-files-memory.md)
