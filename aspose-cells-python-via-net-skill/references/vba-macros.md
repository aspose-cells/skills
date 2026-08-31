# VBA Macros

Read this when reading or adding VBA macros (`.xlsm`), when macros disappear after save, or when you need to sign a VBA project.

Macros live in `workbook.vba_project`, a `aspose.cells.vba.VbaProject`. The modules are `vba_project.modules` and each `VbaModule` exposes `codes` (the VBA source) and `name`. Module types are `aspose.cells.vba.VbaModuleType`: `PROCEDURAL` for a standard module, `CLASS` for a class module, `DOCUMENT`, and `DESIGNER`.

## Pitfalls

### Macros gone after save

Symptom: you added macros but the saved file has none / Excel shows a repair prompt.
Cause: you saved as `SaveFormat.XLSX` (or another non-macro format). Macros are only stored in macro-enabled formats.
Fix: save as `SaveFormat.XLSM` (or `XLSM`). Saving a `.xlsm` out to `.xlsx` also strips macros - pick the format by the file you write, not the file you read.

### Cannot create a VbaProject from scratch in Python

Symptom: `vba.VbaProject()` raises `TypeError: type has no available constructor`.
Cause: the `VbaProject` type has no public constructor in Aspose.Cells for Python via .NET; you cannot attach a brand-new project to a `Workbook`.
Fix: start from a macro-enabled file. Either open an existing `.xlsm` (even an empty macro workbook saved once from Excel) as your template and add modules to its `vba_project`, or save a `.xlsm` once and reuse it. You can still add/replace/delete modules on that project freely.

### Module "type" mismatch

Symptom: `VbaModuleType.MODULE` does not exist.
Cause: Aspose.Cells exposes `PROCEDURAL` (standard module), `CLASS`, `DOCUMENT`, and `DESIGNER` - there is no `MODULE` member. In this Python binding, `add(VbaModuleType)` also fails (`TypeError: can't build Worksheet value from 'VbaModuleType'`), so `add(ws)` is used instead and always creates a `DOCUMENT` module (the type is read-only).
Fix: use `modules.add(ws)` to add a module; you cannot control the type through this binding.

### vba_project is None

Symptom: `workbook.vba_project` is `None`, so `add`/`modules` raise.
Cause: the loaded file is not macro-enabled (e.g. a plain `.xlsx`).
Fix: open a `.xlsm` so a project is present before touching modules.

## Read macros from an existing workbook

```python
import aspose.cells as gc

workbook = gc.Workbook("macro.xlsm")
project = workbook.vba_project                 # None if not a macro-enabled file
if project is not None:
    for module in project.modules:
        print(module.name)
        print(module.codes)                # the VBA source text
```

## Add a standard module

Open a macro-enabled workbook, then append a module with `modules.add(ws)`.

```python
import aspose.cells as gc

workbook = gc.Workbook("template.xlsm")    # must be macro-enabled
project = workbook.vba_project
ws = workbook.worksheets[0]

# add(ws) creates a module; the add(VbaModuleType) overload is not supported
# in this binding (TypeError: can't build Worksheet value from 'VbaModuleType').
# The returned value is the module index; access the module via modules[idx].
idx = project.modules.add(ws)
module = project.modules[idx]
module.name = "MyModule"
module.codes = (
    "Sub Hello()\n"
    "    MsgBox \"hi\"\n"
    "End Sub\n"
)
workbook.save("out.xlsm", gc.SaveFormat.XLSM)   # XLSM keeps the macros
```

## Add a class module

```python
import aspose.cells as gc

workbook = gc.Workbook("template.xlsm")
project = workbook.vba_project
ws = workbook.worksheets[0]

idx = project.modules.add(ws)
cls = project.modules[idx]
cls.name = "MyClass"
cls.codes = (
    "Public x As Integer\n"
    "Sub Inc()\n"
    "    x = x + 1\n"
    "End Sub\n"
)
workbook.save("out.xlsm", gc.SaveFormat.XLSM)
```

## Sign a VBA project

Sign the project (not the file) with an X.509 certificate. `VbaProject.sign` takes a `aspose.cells.digitalsignatures.DigitalSignature` built from a `.pfx`.

```python
import aspose.cells as gc
import aspose.cells.digitalsignatures as ds
from datetime import datetime

workbook = gc.Workbook("template.xlsm")
project = workbook.vba_project
ws = workbook.worksheets[0]

idx = project.modules.add(ws)
module = project.modules[idx]
module.name = "MyModule"
module.codes = "Sub Hello()\n    MsgBox \"hi\"\nEnd Sub\n"

# DigitalSignature takes (cert_bytes, password, sign_time) — 3 args, not 4
with open("cert.pfx", "rb") as f:
    cert_data = f.read()
sig = ds.DigitalSignature(cert_data, "password", datetime.now())
project.sign(sig)
workbook.save("signed.xlsm", gc.SaveFormat.XLSM)
```

The signer is the VBA project object, so the workbook itself stays unsigned; to also sign the file, see digital-signatures-and-document-properties.md.

## Related
- digital-signatures-and-document-properties.md - signing the workbook file and document properties.
- conversion-and-rendering.md - macro-enabled vs plain save formats.
