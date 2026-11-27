# VBA macros

## Preserve macro-enabled formats

Use XLSM, XLS, or another macro-enabled output supported by the installed version. Saving VBA content to XLSX drops macros. Keep the matching save format explicit rather than relying only on the filename.

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook("input.xlsm");
console.log(workbook.getVbaProject().getModules().getCount());
workbook.save("output.xlsm", cells.SaveFormat.XLSM);
```

Aspose.Cells can preserve and manipulate VBA projects; it does not execute VBA macros. A calculated or saved workbook therefore does not run `Workbook_Open`, event handlers, or user-defined VBA functions.

## Add a procedural module

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook();
const project = workbook.getVbaProject();
const index = project.getModules().add(cells.VbaModuleType.PROCEDURAL, "Module1");
const module = project.getModules().get(index);
module.setCodes('Sub Hello()\n  MsgBox "Hello from Aspose.Cells"\nEnd Sub');
workbook.save("macro.xlsm", cells.SaveFormat.XLSM);
```

Obtain `VbaProject` from a workbook; do not try to instantiate it independently. Check for an existing module name before adding one when editing an existing project.

## Signing and security

Sign the VBA project only after all module edits. Subsequent project changes invalidate its signature. Certificate bytes and signing timestamps must use the exact Java-compatible types required by the installed binding; do not assume `Buffer` or JavaScript `Date` conversions. See [signatures](digital-signatures-and-document-properties.md).

Macro-enabled files are active content. Treat uploaded XLSM/XLS files as untrusted, avoid automatic macro execution in downstream consumers, and preserve the original only when that is an explicit requirement.
