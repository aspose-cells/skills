# Digital signatures and document properties

## Inspect existing signatures

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook("signed.xlsx");
const signatures = workbook.getDigitalSignatures();
for (let i = 0; i < signatures.getCount(); i += 1) {
  const signature = signatures.get(i);
  console.log(signature.isValid());
}
```

A signature covers the signed bytes or signed project state. Any later workbook edit or re-save can invalidate it. Verify the final artifact, not an earlier in-memory version.

## Sign last

The signing sequence is:

1. Load or create the workbook and finish every edit.
2. Recalculate formulas and refresh required chart/pivot data.
3. Load the certificate using a Java-compatible byte representation accepted by the installed `DigitalSignature` constructor.
4. Add the signature to a `DigitalSignatureCollection`, then add it to the workbook with the documented API.
5. Save immediately to the final macro-enabled or non-macro format as appropriate.
6. Reopen that final file and verify its signatures.

For VBA signing, use the workbook VBA project's documented `sign(digitalSignature)` operation after all module changes. Workbook signatures and VBA-project signatures protect different content; adding one does not replace the other.

Do not pass `fs.readFileSync()` output as a certificate argument unless the installed Node via Java documentation explicitly supports that conversion. Java `byte[]`, password strings, and signing-time values are bridge boundaries. Avoid invented `Buffer`, JavaScript `Date`, or native certificate-store shortcuts.

Never hard-code certificate private keys or passwords. Load secrets from the deployment secret store and keep unsigned intermediates separate from the final signed artifact.

## Built-in properties

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook();
const properties = workbook.getBuiltInDocumentProperties();
properties.setAuthor("Reporting Service");
properties.setTitle("Monthly Sales");
workbook.save("properties.xlsx", cells.SaveFormat.XLSX);
```

## Custom properties

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook();
workbook.getCustomDocumentProperties().add("Department", "Finance");
workbook.save("custom-properties.xlsx", cells.SaveFormat.XLSX);
```

Document properties are metadata edits, so add them before digital signing. Check the existing custom-property collection before adding duplicates during repeated processing.
