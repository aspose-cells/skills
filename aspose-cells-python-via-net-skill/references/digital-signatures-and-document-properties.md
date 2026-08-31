# Digital Signatures and Document Properties

Read this when signing a workbook (or adding a VBA signature), verifying an existing signature, or reading/writing built-in and custom document properties (title, author, keywords, custom metadata).

Two separate things are both called "signatures" in Aspose.Cells:

- **Workbook digital signature** (`Workbook.add_digital_signature`) - a normal XAdES/PKCS#7 signature over the whole file. Uses a certificate (.pfx/.cer).
- **VBA project signature** (`VbaProject.certificate`) - signs only the macro code. See vba-macros.md.

## Pitfalls

### Signature invalidated by any later edit

Symptom: you sign the workbook, then save it again (or edit a cell) and the signature no longer validates.
Cause: a digital signature covers the file bytes. Any change after signing - re-save, edit, add a shape - invalidates it.
Fix: sign LAST, after all other edits, in the same save sequence. Never sign early and edit later.

```python
import aspose.cells as gc

# WRONG: edit after signing
wb = gc.Workbook("in.xlsx")
cert = open("cert.pfx", "rb").read()
sig = gc.digitalsignatures.DigitalSignature(cert, "password", "comment", __import__("datetime").datetime.now())
sigs = gc.digitalsignatures.DigitalSignatureCollection()
sigs.add(sig)
wb.add_digital_signature(sigs)
wb.worksheets[0].cells.get("A1").put_value("edit after sign")  # invalidates the signature
wb.save("out.xlsx")
```

```python
import aspose.cells as gc
import aspose.cells.digitalsignatures as ds
from datetime import datetime

# RIGHT: all edits first, sign last
wb = gc.Workbook("in.xlsx")
wb.worksheets[0].cells.get("A1").put_value("edit first")
cert = open("cert.pfx", "rb").read()
sig = ds.DigitalSignature(cert, "password", "comment", datetime.now())
sigs = ds.DigitalSignatureCollection()
sigs.add(sig)
wb.add_digital_signature(sigs)
wb.save("out.xlsx")
```

### Signature uses the wrong certificate or fails to load

Symptom: `add_digital_signature` throws, or the saved file shows no signature.
Cause: a `.pfx` requires its password; a `.cer` alone has no private key and cannot sign. Also, `DigitalSignature` is built from a certificate (`bytes` + password, or an `X509Certificate2`), not from a file path.
Fix: load the `.pfx` bytes (or an `X509Certificate2`) with the matching password, wrap it in a `DigitalSignatureCollection`, then add.

```python
import aspose.cells as gc
import aspose.cells.digitalsignatures as ds
from datetime import datetime

cert = open("cert.pfx", "rb").read()
sig = ds.DigitalSignature(cert, "password", "comment", datetime.now())
sigs = ds.DigitalSignatureCollection()
sigs.add(sig)

wb = gc.Workbook("in.xlsx")
wb.add_digital_signature(sigs)
wb.save("out.xlsx")
```

### Custom property "does not stick" / wrong type on read

Symptom: you add a custom property, re-open the file, and the value is missing or the wrong type.
Cause: `CustomDocumentPropertyCollection.add(name, value)` has typed overloads; adding a `str` then reading as a number throws `TypeError`. Also some viewers do not show custom properties unless they are Excel-compatible types.
Fix: pick the typed `add` overload (`str`, `bool`, `int`, `float`, `datetime`) and read back with `.value` after a type check.

```python
import aspose.cells as gc

wb = gc.Workbook()
props = wb.custom_document_properties
props.add("Department", "Finance")        # str
props.add("Approved", True)               # bool
props.add("ApprovalCount", 3)             # int

prop = props.get("Approved")
if isinstance(prop.value, bool):
    print(prop.value)
wb.save("out.xlsx")
```

## Signing a workbook

`Workbook.add_digital_signature(DigitalSignatureCollection)` adds one or more signatures. Build each `DigitalSignature` from the certificate bytes + password (or an `X509Certificate2`), add to a collection, then add the collection. Verify afterwards with `Workbook.is_digitally_signed`.

```python
import aspose.cells as gc
import aspose.cells.digitalsignatures as ds
from datetime import datetime

wb = gc.Workbook("report.xlsx")
# ... build/edit the report ...
cert = open("cert.pfx", "rb").read()
sig = ds.DigitalSignature(cert, "password", "Signed by Finance", datetime.now())
sigs = ds.DigitalSignatureCollection()
sigs.add(sig)
wb.add_digital_signature(sigs)
wb.save("report-signed.xlsx")
print(wb.is_digitally_signed)   # True
```

`Workbook.remove_digital_signature()` removes all signatures. Read existing signatures with `Workbook.get_digital_signature()` (a `DigitalSignatureCollection`) - iterate to read `comments`, `sign_time`, and whether each is valid.

```python
import aspose.cells as gc
import aspose.cells.digitalsignatures as ds

wb = gc.Workbook("signed.xlsx")
if wb.is_digitally_signed:
    for s in wb.get_digital_signature():
        print(f"{s.comments} @ {s.sign_time} valid={s.is_valid}")
    wb.remove_digital_signature()
wb.save("unsigned.xlsx")
```

## VBA signature

The macro project is signed separately via `VbaProject.certificate` (see vba-macros.md). A workbook can carry both a file signature and a VBA signature; they are independent and both must be checked.

## Built-in document properties

`Workbook.built_in_document_properties` (a `DocumentPropertyCollection`) exposes the standard properties by name (`Title`, `Author`, `Subject`, `Keywords`, `Comments`, `Category`, `Company`, `Manager`, ...). Index by name or position.

```python
import aspose.cells as gc

wb = gc.Workbook()
bp = wb.built_in_document_properties
bp.get("Title").value = "Q3 Report"
bp.get("Author").value = "Finance Team"
bp.get("Keywords").value = "quarterly,forecast"
wb.save("out.xlsx")
```

## Custom document properties

`Workbook.custom_document_properties` (a `CustomDocumentPropertyCollection`) stores arbitrary name/value pairs. Use the typed `add` overloads; read back with `.value`.

```python
import aspose.cells as gc
from datetime import datetime

wb = gc.Workbook()
props = wb.custom_document_properties
props.add("ApprovedBy", "Jane")
props.add("ApprovedOn", datetime(2026, 7, 1))
props.add("Version", 2)

for p in props:
    print(f"{p.name} = {p.value}")
wb.save("out.xlsx")
```

## Document protection vs signatures - distinct concepts

Do not confuse signatures with the protection/encryption features in worksheets-rows-columns.md:

| Feature | API | Effect |
|---|---|---|
| Workbook digital signature | `Workbook.add_digital_signature` | proves authorship/integrity; needs a certificate |
| VBA signature | `VbaProject.certificate` | proves the macro code wasn't tampered |
| Open-password encryption | `Workbook.settings.password` | prevents opening without the password |
| Structure protection | `Workbook.protect(ProtectionType.STRUCTURE, pwd)` | blocks sheet add/delete/rename |

## Related
- vba-macros.md - VBA project signing; editing after signing breaks it.
- worksheets-rows-columns.md - open-password encryption and structure protection.
- setup-and-licensing.md - evaluation mode quirks around signing output.
