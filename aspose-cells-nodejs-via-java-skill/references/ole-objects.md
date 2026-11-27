# OLE objects

OLE objects embed or link another document inside a worksheet. They are not pictures, even when the worksheet displays an icon or preview image.

## Inspect existing objects

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook("embedded.xlsx");
const objects = workbook.getWorksheets().get(0).getOleObjects();
for (let i = 0; i < objects.getCount(); i += 1) {
  const object = objects.get(i);
  console.log(object.getProgId());
}
```

Check the object type/prog ID before accessing embedded content. Linked objects can depend on external paths, and an OLE preview does not prove that the embedded payload is present or usable.

## Adding or replacing embedded content

The Java API uses image and file byte arrays when adding an OLE object and setting its data. Node via Java does not guarantee an implicit conversion from JavaScript `Buffer` to Java `byte[]`.

Before generating embedding code, verify all of the following against the installed release:

- The exact `OleObjectCollection.add()` overload and its anchor, width, height, and preview-image arguments.
- A documented way to construct or pass Java-compatible file bytes and preview-image bytes.
- The correct prog ID or object type for the embedded document.
- Whether the output format and target consumer support the embedded object.

If the bridge cannot supply those byte-array types, use a path-based/documented alternative or a Java helper at the boundary; do not pass `fs.readFileSync()` output and claim it is portable.

## Security and portability

Embedded files can contain macros, executable content, external links, or other active payloads. Treat OLE extraction as untrusted file handling: enforce size/type limits, use safe output names, and never execute the payload.

OLE rendering and activation vary across Excel, LibreOffice, web viewers, and non-Windows hosts. Test the final consumer. A converted PDF may contain only a static preview and cannot preserve interactive OLE behavior.
