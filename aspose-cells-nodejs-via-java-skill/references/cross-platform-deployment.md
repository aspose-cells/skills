# Cross-platform deployment

Aspose.Cells for Node.js via Java depends on both Node.js and Java. Test the exact production operating system, CPU architecture, Node release, Java release, and `aspose.cells` version together.

## Native bridge prerequisites

Install a supported JDK and set `JAVA_HOME`. If npm cannot use a prebuilt Node-Java binary, install the compiler toolchain and `node-gyp` prerequisites for the target platform before installing `aspose.cells`. Do not copy random native `.node`, `.dll`, `.so`, or `.dylib` files between machines.

Keep these checks in the deployment smoke test:

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook();
workbook.getWorksheets().get(0).getCells().get("A1").putValue("bridge ok");
workbook.calculateFormula();
console.log(workbook.getWorksheets().getCount());
```

## Linux and Docker

Install Java, fontconfig, and real fonts in the final runtime image, not only in a build stage. Multi-stage images must also contain every native runtime library loaded by the Node-Java bridge.

For a headless Java process, set the JVM option through the environment before starting Node:

```bash
export JAVA_TOOL_OPTIONS=-Djava.awt.headless=true
```

The exact base-image packages vary. On Debian/Ubuntu images, `fontconfig` and a font package such as `fonts-dejavu-core` are a practical baseline. On Alpine, use the Java and font packages provided by that image and verify that the selected Node-Java dependency supports its C library.

Do not assume that spreadsheet creation alone proves PDF/image rendering works. Include a small PDF or PNG conversion in the deployment test.

## Fonts and layout

Spreadsheet cell text, charts, PDF output, and image output use fonts from the host. Missing fonts cause silent substitution, changed text widths, clipped cells, and different page counts.

- Ship and install the fonts used by the workbook.
- Avoid relying on Windows-only font names in Linux containers.
- Configure Aspose font folders or default fonts with the APIs exposed by the installed release when the system font setup is insufficient.
- Set PDF/image layout options explicitly; do not treat the developer workstation's page count as a stable contract.

## Server use

Run independent workbook jobs in isolated processes or worker contexts. Do not share one mutable `Workbook` across requests. Keep file paths job-specific, close application-level resources, and enforce timeouts and memory limits around large or untrusted files.
