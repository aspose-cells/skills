---
name: ole-objects
description: Embed OLE objects (Word docs, PDFs, Excel sheets, text files) into Aspose.Cells for Java worksheets as icons or live previews. Use when the user wants an embedded file inside a cell, or a linked OLE object that refreshes from a source path.
applies_to: Java
keywords: [OLE, embedded object, OleObject, OleObjectCollection, linked, icon, preview]
---

# OLE Objects

OLE objects hang off `Worksheet.getOleObjects()` (an
`OleObjectCollection` in package `com.aspose.cells.Drawing`). Each
`OleObject` wraps an embedded or linked file: `getObjectSourceFullName()`
carries the source (a path for links, the extension for embedded bytes),
`getLabel()` the visible text, and `getDisplayAsIcon()` decides whether it
shows as an icon (default) or a live preview.

## Imports

```java
import com.aspose.cells.*;
```

## Quick example — embed a file from bytes

`add(...)` creates the shell (row, col, and a 1×1 cell size work for a
single-cell anchor), then `setEmbeddedObject(...)` stores the payload with
its extension. The `byte[]` passed to `add` is the **preview image** (a
small PNG icon), not the file content.

```java
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().get(0);

byte[] icon = new byte[]{(byte)0x89, 0x50, 0x4E, 0x47};   // preview PNG header (truncated)
byte[] txtBytes = "Hello from Aspose".getBytes(java.nio.charset.StandardCharsets.UTF_8);

int idx = ws.getOleObjects().add(2, 1, 2, 1, icon);   // topRow, leftColumn, height, width
OleObject ole = ws.getOleObjects().get(idx);

ole.setEmbeddedObject(false, txtBytes, "txt", false, "note.txt");
ole.setLabel("Note");
ole.setDisplayAsIcon(true);
ole.setWidth(120);
ole.setHeight(120);

wb.save("out.xlsx");
```

`setEmbeddedObject(isLink, data, extension, isAutoSize, fileName)`
writes the bytes (or, for `isLink=true`, remembers the source path).

## Pitfalls

### `add(...)` does not embed the file — the bytes go in `setEmbeddedObject`

Symptom: the file appears as an empty shell; double-clicking it does
nothing; or `setEmbeddedObject` is never called and the object has no
payload.
Cause: the 5-arg `add(topRow, leftColumn, height, width, byte[] imageData)`
creates the shell and stores the preview image only. It does **not**
accept a file path or file bytes.
Fix: call `setEmbeddedObject(false, fileBytes, "docx", false, "name.docx")`
right after `add` to load the real content.

### The preview `imageData` must be a valid image, not an empty array

Symptom: `add(...)` throws at runtime.
Cause: the preview is parsed as an image; an empty `byte[]` or a
zero-length payload fails.
Fix: supply a real PNG/JPEG/GIF byte stream as the preview. A
truncated/header-only array still throws — the wrapper expects a fully
decodable image.

### Embedded object opens the wrong app / shows no preview

Symptom: double-clicking the embedded object opens the wrong program,
or the preview area is blank.
Cause: the bytes were embedded with the wrong (or no) extension, so the
registry lookup fails; or `DisplayAsIcon` is true so there is no live
preview by design.
Fix: pass the real extension (`.docx`, `.pdf`, `.xlsx`, …) to
`setEmbeddedObject`, and keep `setDisplayAsIcon(false)` when you want a
visible preview rather than the icon-and-label rendering.

### Link vs embed confusion

Symptom: the user expects the file to refresh from its source, but the
output keeps a snapshot.
Cause: `setEmbeddedObject(false, ...)` copies the bytes in; the object
does **not** track the source file. Use the 6-arg `add` overload or
`setEmbeddedObject(true, ...)` for a live link.
Fix: for links, either

```java
byte[] icon = new byte[]{ /* preview PNG bytes */ };
int idx = ws.getOleObjects().add(0, 0, 0, 0, icon, "C:/data/model.xlsx");
OleObject ole = ws.getOleObjects().get(idx);
ole.setAutoLoad(true);                                  // refresh on open
ole.setObjectSourceFullName("C:/data/model.xlsx");
ole.setDisplayAsIcon(true);
ole.setLabel("Model");
```

After reload, `ole.isLink()` returns `true`.

### `OleObjectCollection` uses `removeAt` / `clear` — no by-reference `remove`

Symptom: `ws.getOleObjects().remove(ole)` does not compile.
Cause: the collection exposes `removeAt(int)` and `clear()` only.
Fix: index-remove, then save.

```java
Workbook wb = new Workbook("has-ole.xlsx");
Worksheet ws = wb.getWorksheets().get(0);

if (ws.getOleObjects().getCount() > 0) {
    ws.getOleObjects().removeAt(0);
}
ws.getOleObjects().clear();
wb.save("out.xlsx");
```

## Reading existing OLE objects

Iterate `Worksheet.getOleObjects()`. `getObjectSourceFullName()` gives
the source (path for links, extension for embedded bytes), `isLink()`
tells you whether the object is linked vs embedded, and
`getDisplayAsIcon()` / `getLabel()` describe how it renders.

```java
Workbook wb = new Workbook("has-ole.xlsx");
Worksheet ws = wb.getWorksheets().get(0);

for (int i = 0; i < ws.getOleObjects().getCount(); i++) {
    OleObject ole = ws.getOleObjects().get(i);
    System.out.println(ole.getLabel() + " from "
        + ole.getObjectSourceFullName() + ", linked=" + ole.isLink());
}
```

## Size, position, and icon

`OleObject` extends `Shape`, so it exposes the same geometry members:
`getWidth()` / `setWidth(int)`, `getHeight()` / `setHeight(int)`,
`setPlacement(int)`, plus `setDisplayAsIcon(boolean)` and `setLabel(...)`
for the icon representation.

```java
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().get(0);

byte[] icon = new byte[]{ /* preview PNG bytes */ };
int idx = ws.getOleObjects().add(5, 1, 5, 1, icon);
OleObject ole = ws.getOleObjects().get(idx);
ole.setEmbeddedObject(false, new byte[]{ /* file bytes */ }, "docx", false, "plan.docx");
ole.setDisplayAsIcon(true);
ole.setLabel("Plan");
ole.setWidth(120);
ole.setHeight(120);
ole.setPlacement(PlacementType.FREE_FLOATING);
wb.save("out.xlsx");
```

## Notes

- `setLinkedCell("A1")` does not stick through save/reload in
  26.x; the reloaded object reports an empty linked cell. Do not rely on
  it.
- OLE objects render into PDF/image output via the standard
  `Workbook.save(path, SaveFormat.Pdf)` path — see
  [../rendering/convert-to-pdf.md](../rendering/convert-to-pdf.md).

## Related

- [shapes-textboxes.md](shapes-textboxes.md) — text boxes, generic
  shapes, and a brief OLE quick-start.
- [insert-image.md](insert-image.md) — pictures, which share the same
  geometry members as `OleObject`.
- [../rendering/convert-to-pdf.md](../rendering/convert-to-pdf.md) —
  OLE objects render into PDF/image output.
