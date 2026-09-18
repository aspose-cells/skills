---
name: digital-signature
description: Sign and verify XLSX with an X.509 digital signature using Aspose.Cells for Java DigitalSignature APIs. Use when the user wants to make a workbook tamper-evident.
applies_to: Java
keywords: [digital signature, X.509, sign, verify, PKCS7, DigitalSignatureCollection]
---

# Digital signatures

## Imports

```java
import com.aspose.cells.*;
```

## Quick example — sign a workbook

```java
Workbook wb = new Workbook("report.xlsx");

// Load a PKCS#12 keystore from a file (commonly `keystore.p12`):
java.security.KeyStore pkcs12 = java.security.KeyStore.getInstance("PKCS12");
try (java.io.FileInputStream fis = new java.io.FileInputStream("keystore.p12")) {
    pkcs12.load(fis, "keystorePassword".toCharArray());
}

DigitalSignatureCollection ds = new DigitalSignatureCollection();
DigitalSignature sig = new DigitalSignature(
        pkcs12, "keystorePassword", "Comments and reason",
        DateTime.getNow());
ds.add(sig);

wb.setDigitalSignature(ds);
wb.save("report-signed.xlsx");
```

## Verify

```java
Workbook wb = new Workbook("report-signed.xlsx");
DigitalSignatureCollection verified = wb.getDigitalSignature();
for (java.util.Iterator<?> it = verified.iterator(); it.hasNext(); ) {
    DigitalSignature s = (DigitalSignature) it.next();
    System.out.println("Comments: " + s.getComments());
    System.out.println("Valid: " + s.isValid());
    java.security.KeyStore ks = s.getCertificate();   // already a KeyStore — no cast needed
    System.out.println("Provider: " + s.getProviderId());
}
```

## Pitfalls

- **Need a `.pfx` (PKCS#12)** containing the private key + cert
  chain. Java's `KeyStore.getInstance("PKCS12")` can produce one
  with `openssl` or `keytool -genkeypair -storetype PKCS12`.
- **`setDigitalSignature` must be called before save** — calling
  save re-flattens the zip; signing afterwards is a no-op.
- **The signature is embedded inside the OOXML package**, not
  detached — recipients see it under File → Info → View Signatures
  in Excel.
- **`isValid()` returns the cached value** for the *first* signature
  in the collection at the moment `load` happened. To re-verify
  after edits, run a fresh validation.
- **BouncyCastle** jars are suggested for any project, and required
  for VBA project signatures (`VbaProject.sign(...)`). Aspose.Cells
  registers the provider itself.

## Related

- [encryption-decryption.md](encryption-decryption.md)
- [../getting-started/maven-gradle-dependencies.md](../getting-started/maven-gradle-dependencies.md)
