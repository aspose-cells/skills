# Setup and licensing

## Install the correct binding

```bash
npm install aspose.cells@26.9.0
```

Use CommonJS and the JavaBean API:

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook();
workbook.getWorksheets().get(0).getCells().get("A1").putValue("Hello");
workbook.save("hello.xlsx", cells.SaveFormat.XLSX);
```

Do not install or mix in `aspose.cells.node`; that is Aspose.Cells for Node.js via C++ and has a different API. The Java binding also needs a working Java runtime/JDK and, when no compatible prebuilt bridge is available, the native build prerequisites of its Node-Java dependency.

## Apply a license once

Set the license before creating any `Workbook`. Use an absolute path in services and CLI tools so the result does not depend on the process working directory.

```javascript
const path = require("node:path");
const cells = require("aspose.cells");

const license = new cells.License();
license.setLicense(path.resolve("licenses/Aspose.Cells.lic"));

const workbook = new cells.Workbook();
if (!workbook.isLicensed()) {
  throw new Error("Aspose.Cells is running in evaluation mode");
}
```

Keep license initialization in one startup module. Do not catch a licensing error and continue silently: evaluation output can contain watermarks, extra evaluation sheets, or other limitations. A package released after a subscription's support end date can also run in evaluation mode even when the license file itself is valid.

For an intentional evaluation run, omit `setLicense()` and state that the output is evaluation output. Never copy a license into source control or embed its contents in generated code.

## Diagnose startup failures

- `Cannot find module 'aspose.cells'`: install in the application project, not this skill directory.
- Java or bridge load errors: verify `JAVA_HOME`, architecture (x64 versus arm64), Node version, and native build prerequisites.
- `isLicensed()` is false: check the path, license product, expiration/subscription coverage, and package version.
- `require("aspose.cells.node")`: wrong product binding; replace it with `require("aspose.cells")` and JavaBean calls.
