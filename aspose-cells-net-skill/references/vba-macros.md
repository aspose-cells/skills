# VBA Macros

Read this when creating, reading, modifying, or signing VBA macros in an `.xlsm`/`.xlsb` workbook, converting an `.xls` to `.xlsm`, or when a macro-enabled file loses its code or its signature.

Macros only survive in macro-enabled formats (`.xlsm`, `.xlsb`, `.xls`). Saving to plain `.xlsx` silently drops all VBA. The macro project is exposed as `Workbook.VbaProject` (a `VbaProject`). It is **not** null on a freshly created workbook - `new Workbook()` already carries a project with the built-in `ThisWorkbook` and `Sheet1` document modules. There is no public `VbaProject(...)` constructor and `Workbook.VbaProject` is read-only, so you never create or assign the project yourself: you add modules to it.

## Pitfalls

### Macros disappear when saving to .xlsx

Symptom: the output file has no macros, even though the input did.
Cause: `.xlsx` cannot carry a VBA project. Aspose strips it silently.
Fix: save macro work as `SaveFormat.Xlsm` (or `Xlsb`), never `Xlsx`. Keep `.xlsx` for macro-free deliverables.

```csharp
var workbook = new Workbook("with-macros.xlsm");
workbook.Save("out.xlsm", SaveFormat.Xlsm);   // macros preserved
workbook.Save("out.xlsx", SaveFormat.Xlsx);   // macros dropped - don't
```

### "VbaProject is null / must construct it" confusion

Symptom: you copied old sample code that calls `new VbaProject(...)` or assigns `workbook.VbaProject = ...` and it does not compile.
Cause: those patterns do not exist in 26.x. The `VbaProject` constructor is internal and `Workbook.VbaProject` has no setter.
Fix: use `workbook.VbaProject` directly and call `Modules.Add(...)` on it.

```csharp
var workbook = new Workbook();
VbaProject vba = workbook.VbaProject;   // never null

int idx = vba.Modules.Add(VbaModuleType.Procedural, "MyModule");
VbaModule module = vba.Modules[idx];
module.Codes = "Sub Hello()\r\n    MsgBox \"Hello from Aspose\"\r\nEnd Sub";
workbook.Save("out.xlsm", SaveFormat.Xlsm);
```

### Modules appear but the code does not run / signature breaks

Symptom: `.xlsm` opens but macros are disabled or the code is gone; or you signed a workbook and the signature no longer validates after a re-save.
Cause: adding code via `module.Codes` writes plain text but does not produce a valid signed project; digital signatures are invalidated by any edit to the signed content (including a plain re-save after editing).
Fix: sign *last*, after all macro edits, with `VbaProject.Sign(...)` or `Workbook.AddDigitalSignature(...)` (see digital-signatures-and-document-properties.md). Do not edit macros after signing.

## Reading existing macros

Iterate the project's modules and read `VbaModule.Codes` (the raw VBA source). `VbaModule.Type` tells you the module kind (Procedural, Document, Class, Designer).

```csharp
var workbook = new Workbook("with-macros.xlsm");
VbaProject vba = workbook.VbaProject;
if (vba.Modules.Count == 0) { Console.WriteLine("no modules"); return; }

foreach (VbaModule module in vba.Modules)
    Console.WriteLine($"{module.Name} [{module.Type}] -> {module.Codes}");
```

## Adding a new module with code

`Modules.Add(VbaModuleType.Procedural, name)` returns the index; `VbaModuleType` is `Procedural`, `Document`, `Class`, or `Designer`. Set `Codes` to the full VBA source text.

```csharp
var workbook = new Workbook();
VbaProject vba = workbook.VbaProject;

int idx = vba.Modules.Add(VbaModuleType.Procedural, "Hello");
VbaModule module = vba.Modules[idx];
module.Codes =
    "Public Function Add(a As Double, b As Double) As Double\r\n" +
    "    Add = a + b\r\n" +
    "End Function";
workbook.Save("out.xlsm", SaveFormat.Xlsm);
```

## Project-level properties and signing

`VbaProject` exposes `IsSigned`, `Name`, `Encoding`, `IsProtected`, and `References`. Sign the macro project with `vba.Sign(DigitalSignature)` (same `DigitalSignature` type used for workbook signatures - see digital-signatures-and-document-properties.md). Protect the VBA password with `vba.Protect(bool lockForViewing, string password)`; verify later with `ValidatePassword`.

```csharp
var workbook = new Workbook("with-macros.xlsm");
VbaProject vba = workbook.VbaProject;

foreach (VbaProjectReference reference in vba.References)
    Console.WriteLine(reference.Name);

if (!vba.IsSigned)
{
    var sig = new DigitalSignature(
        File.ReadAllBytes("cert.pfx"), "password", "macro signing", DateTime.Now);
    vba.Sign(sig);   // VbaProject.Sign takes ONE DigitalSignature (not a collection)
}
workbook.Save("out.xlsm", SaveFormat.Xlsm);
```

## Converting .xls to .xlsm to keep macros

Legacy `.xls` can carry VBA; converting to `.xlsx` drops it. Convert to `.xlsm` instead when macros must survive a modernization.

```csharp
var workbook = new Workbook("legacy.xls");
if (workbook.VbaProject.Modules.Count > 0)
{
    workbook.Save("modernized.xlsm", SaveFormat.Xlsm);
}
```

## Related
- digital-signatures-and-document-properties.md - signing a workbook; a signature invalidates later edits.
- setup-and-licensing.md - evaluation mode limits macro-related output.
- worksheets-rows-columns.md - Xlsm save format limits (1,048,576 x 16,384).
