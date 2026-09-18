---
name: vba-macros
description: Create, read, modify, and sign VBA macros in .xlsm/.xlsb/.xls workbooks with Aspose.Cells for Java. Use when the user wants to add a macro module, preserve macros across save, modernize legacy .xls, or sign the macro project with a certificate.
applies_to: Java
keywords: [VBA, macros, VbaProject, VbaModule, xlsm, xlsb, digital signature, macro project]
---

# VBA macros

Macros only survive in macro-enabled formats (`.xlsm`, `.xlsb`, `.xls`).
Saving to plain `.xlsx` silently drops all VBA. The macro project is
exposed as `Workbook.getVbaProject()` (a `VbaProject`). It is **not**
null on a freshly created workbook — `new Workbook()` already carries
a project with the built-in ThisWorkbook Excel VBA document module plus
a per-worksheet document module for each sheet. There is no public
`VbaProject` constructor and the `VbaProject` property on `Workbook`
has no setter, so you never create or assign the project yourself:
you add modules to it.

## Imports

```java
import com.aspose.cells.*;
```

## Pitfalls

### Macros disappear when saving to `.xlsx`

Symptom: the output file has no macros, even though the input did.
Cause: `.xlsx` cannot carry a VBA project. Aspose strips it silently.
Fix: save macro work as `SaveFormat.XLSM` (or `XLSB`), never `XLSX`.
Keep `.xlsx` for macro-free deliverables.

```java
Workbook wb = new Workbook("with-macros.xlsm");
wb.save("out.xlsm", SaveFormat.XLSM);   // macros preserved
wb.save("out.xlsx", SaveFormat.XLSX);   // macros dropped — don't
```

### "VbaProject is null / must construct it" confusion

Symptom: code that constructs a `VbaProject` directly, or that assigns a
new project to the `setVbaProject` setter on `Workbook`, does not compile.
Cause: those patterns do not exist. `VbaProject` has no public
constructor and the `setVbaProject` setter does not exist on
`Workbook`.
Fix: use `wb.getVbaProject()` directly and call
`getModules().add(VbaModuleType.PROCEDURAL, "ModuleName")` on it.

```java
Workbook wb = new Workbook();
VbaProject vba = wb.getVbaProject();   // never null

int idx = vba.getModules().add(VbaModuleType.PROCEDURAL, "MyModule");
VbaModule module = vba.getModules().get(idx);
module.setCodes(
    "Sub Hello()\r\n" +
    "    MsgBox \"Hello from Aspose\"\r\n" +
    "End Sub\r\n");
wb.save("out.xlsm", SaveFormat.XLSM);
```

### Modules appear but the code does not run / signature breaks

Symptom: `.xlsm` opens but macros are disabled or the code is gone; or
you signed the workbook and the signature no longer validates after a
re-save.
Cause: any edit to the signed content (including a plain re-save after
editing) invalidates the signature.
Fix: sign **last**, after all macro edits, in the same save sequence.

## Reading existing macros

Iterate the project's modules and read `VbaModule.getCodes()` (the raw
VBA source). `VbaModule.getType()` reports the module kind
(`VbaModuleType.PROCEDURAL`, `DOCUMENT`, `CLASS`, `DESIGNER`).

```java
Workbook wb = new Workbook("with-macros.xlsm");
VbaProject vba = wb.getVbaProject();
if (vba.getModules().getCount() == 0) {
    System.out.println("no modules");
    return;
}

for (int i = 0; i < vba.getModules().getCount(); i++) {
    VbaModule module = vba.getModules().get(i);
    System.out.println(module.getName() + " [" + module.getType()
        + "] -> " + module.getCodes());
}
```

## Adding a new module with code

`getModules().add(VbaModuleType.PROCEDURAL, name)` returns the index;
`VbaModuleType` is `PROCEDURAL`, `DOCUMENT`, `CLASS`, or `DESIGNER`.
Set `setCodes(String)` to the full VBA source text.

```java
Workbook wb = new Workbook();
VbaProject vba = wb.getVbaProject();

int idx = vba.getModules().add(VbaModuleType.PROCEDURAL, "Hello");
VbaModule module = vba.getModules().get(idx);
module.setCodes(
    "Public Function Add(a As Double, b As Double) As Double\r\n" +
    "    Add = a + b\r\n" +
    "End Function\r\n");
wb.save("out.xlsm", SaveFormat.XLSM);
```

To add a per-worksheet document module (the per-sheet `ThisWorkbook`-style
module Excel creates for each worksheet):

```java
Workbook wb = new Workbook();
VbaProject vba = wb.getVbaProject();

int docIdx = vba.getModules().add(VbaModuleType.DOCUMENT, "Sheet1");
VbaModule docModule = vba.getModules().get(docIdx);
docModule.setCodes(
    "Private Sub Worksheet_Activate()\r\n" +
    "    MsgBox \"Sheet1 activated\"\r\n" +
    "End Sub\r\n");
wb.save("out.xlsm", SaveFormat.XLSM);
```

`VbaModuleCollection.add(int, String)` is also available — pass the
worksheet index (0-based) and a module name to register a document
module against an existing sheet.

## Project-level properties and signing

`VbaProject` exposes `isSigned()`, `getName()` / `setName(String)`,
`getEncoding()` / `setEncoding(Encoding)`, `isProtected()`, and
`getReferences()` (a `VbaProjectReferenceCollection`). Sign the macro
project with `vba.sign(DigitalSignature)` — the same `DigitalSignature`
type used for workbook signatures; see
[../security/digital-signature.md](../security/digital-signature.md).
Protect the VBA with `vba.protect(boolean lockForViewing, String password)`
and verify later with `vba.validatePassword(String)`.

```java
Workbook wb = new Workbook("with-macros.xlsm");
VbaProject vba = wb.getVbaProject();

for (int i = 0; i < vba.getReferences().getCount(); i++) {
    System.out.println(vba.getReferences().get(i).getName());
}

if (!vba.isSigned()) {
    java.security.KeyStore pkcs12 = java.security.KeyStore.getInstance("PKCS12");
    try (java.io.FileInputStream fis = new java.io.FileInputStream("cert.pfx")) {
        pkcs12.load(fis, "password".toCharArray());
    }
    DigitalSignature sig = new DigitalSignature(
        pkcs12, "password", "macro signing", DateTime.getNow());
    vba.sign(sig);
}
wb.save("out.xlsm", SaveFormat.XLSM);
```

## Converting `.xls` to `.xlsm` to keep macros

Legacy `.xls` can carry VBA; converting to `.xlsx` drops it. Convert to
`.xlsm` instead when macros must survive a modernization.

```java
Workbook wb = new Workbook("legacy.xls");
if (wb.getVbaProject().getModules().getCount() > 0) {
    wb.save("modernized.xlsm", SaveFormat.XLSM);
}
```

## Removing or clearing modules

`VbaModuleCollection` exposes `remove(String name)`,
`remove(Worksheet)`, `removeAt(int)`, and `clear()`. Use these when
you want to strip a project down to specific modules.

```java
Workbook wb = new Workbook("with-macros.xlsm");
VbaProject vba = wb.getVbaProject();

vba.getModules().remove("Hello");        // by name
wb.save("out.xlsm", SaveFormat.XLSM);
```

## Notes

- `xlsm` save format limits: 1,048,576 rows × 16,384 columns (same as
  `.xlsx`). For larger or binary-only needs, use `.xlsb` (binary) via
  `SaveFormat.XLSB`.
- The signature on a VBA project is independent of any workbook-level
  signature. A workbook can carry both, and both must be verified
  separately.
- `isValidSigned()` returns whether the certificate chain on the
  embedded signature still verifies — useful for a quick
  trust check.

## Related

- [../security/digital-signature.md](../security/digital-signature.md) —
  workbook-level signing; a signature invalidates later edits.
- [../security/document-properties.md](../security/document-properties.md) —
  built-in and custom document metadata.
- [../getting-started/save-export.md](../getting-started/save-export.md) —
  save-format selection.
