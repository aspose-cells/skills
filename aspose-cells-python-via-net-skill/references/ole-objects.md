# OLE Objects and Embedded Content

Read this when embedding another file (Word doc, PDF, Excel sheet, text) inside a worksheet, when setting the OLE icon and label, or when the embedded object opens the wrong app / shows no preview.

OLE objects hang off `worksheet.ole_objects` (an `OleObjectCollection`, namespace `aspose.cells.drawing`). Each `OleObject` wraps an embedded or linked file: `object_source_full_name` carries the source (a path for links, the extension for embedded bytes), `label` the visible text, and `display_as_icon` decides whether it shows as an icon (default) or as a live preview.

Two facts that look like bugs but are the real API in 26.x:

- `OleObjectCollection.add` takes **cell-rect coordinates** plus a preview `image_data`, NOT a file path plus `bool is_linked`. Signatures:
  - `add(int top_row, int left_column, int height, int width, bytes image_data)` - embeds an empty shell (see `set_embedded_object` below to put real bytes in).
  - `add(int top_row, int left_column, int height, int width, bytes image_data, str linked_file)` - a **linked** object; `linked_file` is the source path and the object refreshes from it.
  The `image_data` parameter is the preview image (e.g. a small PNG icon); it is parsed as an image, so a zero-length array raises. There is no `add(row, col, bytes data, str extension)` overload.
- To actually embed file **bytes**, create the shell with `add`, then call `OleObject.set_embedded_object(bool is_link, bytes image_data, str extension, bool is_auto_size, str file_name)`. The 5-arg `add` alone stores an object with no payload.

## Pitfalls

### Embedded object shows the wrong content / opens the wrong app

Symptom: double-clicking the embedded object opens the wrong program, or the preview is blank.
Cause: the bytes were embedded with the wrong (or no) extension, so the registry lookup fails; or `display_as_icon` is True so there is no preview by design.
Fix: pass the real extension to `set_embedded_object` and keep `display_as_icon = False` when you want a visible preview.

### Link vs embed confusion

Symptom: the file must always be up to date, but you embedded a snapshot.
Cause: `set_embedded_object` with `is_link = False` copies the bytes in; the object does not track the source file.
Fix: use the 6-arg `add(..., linked_file)` overload (or `set_embedded_object(True, ...)`) when the object must refresh from its source on open. Linked objects still need the source reachable.

### "Add throws" - wrong overload or empty image_data

Symptom: `sheet.ole_objects.add(...)` throws.
Cause: the old sample `add(row, col, bytes data, str extension)` does not exist in 26.x, and `add`'s `image_data` must be a valid image. Passing an empty array or a wrong overload fails at runtime.
Fix: use the 5/6-arg rectangle overloads above and call `set_embedded_object` to load real content.

## Embed a file from bytes

`add` creates the shell (row, col, and 1x1 cell size work for a single-cell anchor), then `set_embedded_object` stores the payload with its extension.

```python
import aspose.cells as gc
import aspose.cells.drawing as dw

workbook = gc.Workbook()
sheet = workbook.worksheets[0]

with open("C:/temp/preview.png", "rb") as f:
    icon = f.read()                 # preview image, any size
with open("C:/temp/note.txt", "rb") as f:
    txt_bytes = f.read()            # the embedded file

idx = sheet.ole_objects.add(2, 1, 2, 1, icon)   # top_row, left_column, height, width
ole = sheet.ole_objects[idx]

ole.set_embedded_object(False, txt_bytes, "txt", False, "note.txt")  # embed bytes
ole.label = "Note"
ole.display_as_icon = True       # or False for a live preview
ole.width = 120
ole.height = 120
workbook.save("out.xlsx", gc.SaveFormat.XLSX)
```

Note: setting `OleObject.linked_cell = "A1"` does not stick through a save/reload in 26.x - the reloaded object reports an empty `linked_cell`. Do not rely on it for OLE objects.

## Embed a file as a link

The 6-arg `add` overload keeps the object tied to the source file. `auto_load` refreshes from the source when the workbook opens.

```python
import aspose.cells as gc
import aspose.cells.drawing as dw

workbook = gc.Workbook()
sheet = workbook.worksheets[0]

with open("C:/temp/preview.png", "rb") as f:
    icon = f.read()
idx = sheet.ole_objects.add(0, 0, 0, 0, icon, "C:/data/model.xlsx")
ole = sheet.ole_objects[idx]

ole.auto_load = True             # refresh from source when the workbook opens
ole.object_source_full_name = "C:/data/model.xlsx"
ole.display_as_icon = True
ole.label = "Model"
workbook.save("out.xlsx", gc.SaveFormat.XLSX)
```

After reload, `ole.is_link` is `True` and `object_source_full_name` holds the source path.

## Size, position, and icon

`OleObject` exposes the same geometry members as pictures: `width`, `height`, `top`, `left`, `placement`, plus `display_as_icon` and `label` for the icon representation.

```python
import aspose.cells as gc
import aspose.cells.drawing as dw

workbook = gc.Workbook()
sheet = workbook.worksheets[0]

with open("C:/temp/preview.png", "rb") as f:
    icon = f.read()
idx = sheet.ole_objects.add(5, 1, 5, 1, icon)
ole = sheet.ole_objects[idx]
ole.set_embedded_object(False, open("C:/temp/plan.docx", "rb").read(), "docx", False, "plan.docx")
ole.display_as_icon = True
ole.label = "Plan"
ole.width = 120
ole.height = 120
ole.placement = dw.PlacementType.FREE_FLOATING
workbook.save("out.xlsx", gc.SaveFormat.XLSX)
```

## Read existing OLE objects

Iterate `sheet.ole_objects`; `object_source_full_name` gives the source (path for links, extension for embedded bytes), `is_link` tells you whether it is linked (not embedded), and `display_as_icon`/`label` describe how it renders.

```python
import aspose.cells as gc

workbook = gc.Workbook("has-ole.xlsx")
sheet = workbook.worksheets[0]

for ole in sheet.ole_objects:
    print(f"{ole.label} from {ole.object_source_full_name}, linked={ole.is_link}")
```

## Delete an OLE object

`OleObjectCollection` exposes `remove(object)` (pass the `OleObject`, not an index) and `clear()`.

```python
import aspose.cells as gc

workbook = gc.Workbook("has-ole.xlsx")
sheet = workbook.worksheets[0]

if sheet.ole_objects.length > 0:
    sheet.ole_objects.remove(sheet.ole_objects[0])   # pass the OleObject, not an index
sheet.ole_objects.clear()
workbook.save("out.xlsx")
```

## Related
- images-and-shapes.md - pictures share the same geometry members (width/height/top/left/placement).
- conversion-and-rendering.md - OLE objects render into PDF/image output.
