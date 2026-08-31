# Digital Signatures and Document Properties

Read this when signing a workbook (or adding a VBA signature), verifying an existing signature, or reading/writing built-in and custom document properties (title, author, keywords, custom metadata).

Two separate things are both called "signatures" in Aspose.Cells:

- **Workbook digital signature** (`Workbook.AddDigitalSignature`) - a normal XAdES/PKCS#7 signature over the whole file. Uses a certificate (.pfx/.cer).
- **VBA project signature** (`VbaProject.Certificate`) - signs only the macro code. See vba-macros.md.

## Pitfalls

### Signature invalidated by any later edit

Symptom: you sign the workbook, then save it again (or edit a cell) and the signature no longer validates.
Cause: a digital signature covers the file bytes. Any change after signing - re-save, edit, add a shape - invalidates it.
Fix: sign LAST, after all other edits, in the same save sequence. Never sign early and edit later.

```csharp
// WRONG: edit after signing
Workbook wb = new Workbook("in.xlsx");
var sig = new DigitalSignature(File.ReadAllBytes("cert.pfx"), "password", "comment", DateTime.Now);
var sigs = new DigitalSignatureCollection();
sigs.Add(sig);
wb.AddDigitalSignature(sigs);
wb.Worksheets[0].Cells["A1"].PutValue("edit after sign"); // invalidates the signature
wb.Save("out.xlsx");
```

```csharp
// RIGHT: all edits first, sign last
Workbook wb = new Workbook("in.xlsx");
wb.Worksheets[0].Cells["A1"].PutValue("edit first");
var sig = new DigitalSignature(File.ReadAllBytes("cert.pfx"), "password", "comment", DateTime.Now);
var sigs = new DigitalSignatureCollection();
sigs.Add(sig);
wb.AddDigitalSignature(sigs);
wb.Save("out.xlsx");
```

### Signature uses the wrong certificate or fails to load

Symptom: `AddDigitalSignature` throws, or the saved file shows no signature.
Cause: a `.pfx` requires its password; a `.cer` alone has no private key and cannot sign. Also, `DigitalSignature` is built from a certificate (`byte[]` + password, or an `X509Certificate2`), not from a file path - the old `AddDigitalSignature(DigitalSignature)` single-signature overload was replaced by a collection overload in 26.x.
Fix: load the `.pfx` bytes (or an `X509Certificate2`) with the matching password, wrap it in a `DigitalSignatureCollection`, then add.

```csharp
byte[] cert = File.ReadAllBytes("cert.pfx");
var sig = new DigitalSignature(cert, "password", "comment", DateTime.Now);
var sigs = new DigitalSignatureCollection();
sigs.Add(sig);

Workbook wb = new Workbook("in.xlsx");
wb.AddDigitalSignature(sigs);
wb.Save("out.xlsx");
```

### Custom property "does not stick" / wrong type on read

Symptom: you add a custom property, re-open the file, and the value is missing or the wrong type.
Cause: `CustomDocumentPropertyCollection.Add(name, value)` has typed overloads; adding a `string` then reading as a number throws `InvalidCastException`. Also some viewers do not show custom properties unless they are Excel-compatible types.
Fix: pick the typed `Add` overload (`string`, `bool`, `int`, `double`, `DateTime`) and read back with `.Value` (an `object`) after a type check.

```csharp
Workbook wb = new Workbook();
CustomDocumentPropertyCollection props = wb.CustomDocumentProperties;
props.Add("Department", "Finance");        // string
props.Add("Approved", true);               // bool
props.Add("ApprovalCount", 3);             // int

DocumentProperty prop = props["Approved"];
if (prop.Value is bool approved) Console.WriteLine(approved);
wb.Save("out.xlsx");
```

## Signing a workbook

`Workbook.AddDigitalSignature(DigitalSignatureCollection)` adds one or more signatures. Build each `DigitalSignature` from the certificate bytes + password (or an `X509Certificate2`), add to a collection, then add the collection. Verify afterwards with `Workbook.IsDigitallySigned`.

```csharp
Workbook wb = new Workbook("report.xlsx");
// ... build/edit the report ...
var sig = new DigitalSignature(
    File.ReadAllBytes("cert.pfx"), "password", "Signed by Finance", DateTime.Now);
var sigs = new DigitalSignatureCollection();
sigs.Add(sig);
wb.AddDigitalSignature(sigs);
wb.Save("report-signed.xlsx");
Console.WriteLine(wb.IsDigitallySigned);   // True
```

`Workbook.RemoveDigitalSignature()` removes all signatures. Read existing signatures with `Workbook.GetDigitalSignature()` (a `DigitalSignatureCollection`) - iterate to read `Comments`, `SignTime`, and whether each is valid.

```csharp
Workbook wb = new Workbook("signed.xlsx");
if (wb.IsDigitallySigned)
{
    foreach (DigitalSignature s in wb.GetDigitalSignature())
        Console.WriteLine($"{s.Comments} @ {s.SignTime} valid={s.IsValid}");
    wb.RemoveDigitalSignature();
}
wb.Save("unsigned.xlsx");
```

## VBA signature

The macro project is signed separately via `VbaProject.Certificate` (see vba-macros.md). A workbook can carry both a file signature and a VBA signature; they are independent and both must be checked.

## Built-in document properties

`Workbook.BuiltInDocumentProperties` (a `DocumentPropertyCollection`) exposes the standard properties by name (`Title`, `Author`, `Subject`, `Keywords`, `Comments`, `Category`, `Company`, `Manager`, ...). Index by name or position.

```csharp
Workbook wb = new Workbook();
DocumentPropertyCollection bp = wb.BuiltInDocumentProperties;
bp["Title"].Value = "Q3 Report";
bp["Author"].Value = "Finance Team";
bp["Keywords"].Value = "quarterly,forecast";
wb.Save("out.xlsx");
```

## Custom document properties

`Workbook.CustomDocumentProperties` (a `CustomDocumentPropertyCollection`) stores arbitrary name/value pairs. Use the typed `Add` overloads; read back with `.Value`.

```csharp
Workbook wb = new Workbook();
CustomDocumentPropertyCollection props = wb.CustomDocumentProperties;
props.Add("ApprovedBy", "Jane");
props.Add("ApprovedOn", new DateTime(2026, 7, 1));
props.Add("Version", 2);

foreach (DocumentProperty p in props)
    Console.WriteLine($"{p.Name} = {p.Value}");
wb.Save("out.xlsx");
```

## Document protection vs signatures - distinct concepts

Do not confuse signatures with the protection/encryption features in worksheets-rows-columns.md:

| Feature | API | Effect |
|---|---|---|
| Workbook digital signature | `Workbook.AddDigitalSignature` | proves authorship/integrity; needs a certificate |
| VBA signature | `VbaProject.Certificate` | proves the macro code wasn't tampered |
| Open-password encryption | `Workbook.Settings.Password` | prevents opening without the password |
| Structure protection | `Workbook.Protect(ProtectionType.Structure, pwd)` | blocks sheet add/delete/rename |

## Related
- vba-macros.md - VBA project signing; editing after signing breaks it.
- worksheets-rows-columns.md - open-password encryption and structure protection.
- setup-and-licensing.md - evaluation mode quirks around signing output.
