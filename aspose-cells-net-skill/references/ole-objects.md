# OLE Objects and Embedded Content

Read this when embedding another file (Word doc, PDF, Excel sheet, text) inside a worksheet, when setting the OLE icon and label, or when the embedded object opens the wrong app / shows no preview.

OLE objects hang off `worksheet.OleObjects` (an `OleObjectCollection`, namespace `Aspose.Cells.Drawing`). Each `OleObject` wraps an embedded or linked file: `ObjectSourceFullName` carries the source (a path for links, the extension for embedded bytes), `Label` the visible text, and `DisplayAsIcon` decides whether it shows as an icon (default) or as a live preview.

Two facts that look like bugs but are the real API in 26.x:

- `OleObjectCollection.Add` takes **cell-rect coordinates** plus a preview `imageData`, NOT a file path plus `bool isLinked`. Signatures:
  - `Add(int topRow, int leftColumn, int height, int width, byte[] imageData)` - embeds an empty shell (see `SetEmbeddedObject` below to put real bytes in).
  - `Add(int topRow, int leftColumn, int height, int width, byte[] imageData, string linkedFile)` - a **linked** object; `linkedFile` is the source path and the object refreshes from it.
  The `imageData` parameter is the preview image (e.g. a small PNG icon); it is parsed as an image, so a zero-length array throws. There is no `Add(row, col, byte[] data, string extension)` overload.
- To actually embed file **bytes**, create the shell with `Add`, then call `OleObject.SetEmbeddedObject(bool isLink, byte[] imageData, string extension, bool isAutoSize, string fileName)`. The 5-arg `Add` alone stores an object with no payload.

## Pitfalls

### Embedded object shows the wrong content / opens the wrong app

Symptom: double-clicking the embedded object opens the wrong program, or the preview is blank.
Cause: the bytes were embedded with the wrong (or no) extension, so the registry lookup fails; or `DisplayAsIcon` is true so there is no preview by design.
Fix: pass the real extension to `SetEmbeddedObject` and keep `DisplayAsIcon = false` when you want a visible preview.

### Link vs embed confusion

Symptom: the file must always be up to date, but you embedded a snapshot.
Cause: `SetEmbeddedObject` with `isLink = false` copies the bytes in; the object does not track the source file.
Fix: use the 6-arg `Add(..., linkedFile)` overload (or `SetEmbeddedObject(true, ...)`) when the object must refresh from its source on open. Linked objects still need the source reachable.

### "Add throws" - wrong overload or empty imageData

Symptom: `sheet.OleObjects.Add(...)` throws.
Cause: the old sample `Add(row, col, byte[] data, string extension)` does not exist in 26.x, and `Add`'s `imageData` must be a valid image. Passing an empty array or a wrong overload fails at runtime.
Fix: use the 5/6-arg rectangle overloads above and call `SetEmbeddedObject` to load real content.

## Embed a file from bytes

`Add` creates the shell (row, col, and 1x1 cell size work for a single-cell anchor), then `SetEmbeddedObject` stores the payload with its extension.

```csharp
var workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];

byte[] icon = File.ReadAllBytes("C:\\temp\\preview.png");   // preview image, any size
byte[] txtBytes = File.ReadAllBytes("C:\\temp\\note.txt");  // the embedded file

int idx = sheet.OleObjects.Add(2, 1, 2, 1, icon);   // topRow, leftColumn, height, width
OleObject ole = sheet.OleObjects[idx];

ole.SetEmbeddedObject(false, txtBytes, "txt", false, "note.txt");  // embed bytes
ole.Label = "Note";
ole.DisplayAsIcon = true;       // or false for a live preview
ole.Width = 120;
ole.Height = 120;
workbook.Save("out.xlsx", SaveFormat.Xlsx);
```

Note: setting `OleObject.LinkedCell = "A1"` does not stick through a save/reload in 26.x - the reloaded object reports an empty `LinkedCell`. Do not rely on it for OLE objects.

## Embed a file as a link

The 6-arg `Add` overload keeps the object tied to the source file. `AutoLoad` refreshes from the source when the workbook opens.

```csharp
var workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];

byte[] icon = File.ReadAllBytes("C:\\temp\\preview.png");
int idx = sheet.OleObjects.Add(0, 0, 0, 0, icon, "C:\\data\\model.xlsx");
OleObject ole = sheet.OleObjects[idx];

ole.AutoLoad = true;             // refresh from source when the workbook opens
ole.ObjectSourceFullName = "C:\\data\\model.xlsx";
ole.DisplayAsIcon = true;
ole.Label = "Model";
workbook.Save("out.xlsx", SaveFormat.Xlsx);
```

After reload, `ole.IsLink` is `true` and `ObjectSourceFullName` holds the source path.

## Size, position, and icon

`OleObject` exposes the same geometry members as pictures: `Width`, `Height`, `Top`, `Left`, `Placement`, plus `DisplayAsIcon` and `Label` for the icon representation.

```csharp
var workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];

byte[] icon = File.ReadAllBytes("C:\\temp\\preview.png");
int idx = sheet.OleObjects.Add(5, 1, 5, 1, icon);
OleObject ole = sheet.OleObjects[idx];
ole.SetEmbeddedObject(false, File.ReadAllBytes("C:\\temp\\plan.docx"), "docx", false, "plan.docx");
ole.DisplayAsIcon = true;
ole.Label = "Plan";
ole.Width = 120;
ole.Height = 120;
ole.Placement = PlacementType.FreeFloating;
workbook.Save("out.xlsx", SaveFormat.Xlsx);
```

## Read existing OLE objects

Iterate `sheet.OleObjects`; `ObjectSourceFullName` gives the source (path for links, extension for embedded bytes), `IsLink` tells you whether it is linked (not embedded), and `DisplayAsIcon`/`Label` describe how it renders.

```csharp
var workbook = new Workbook("has-ole.xlsx");
Worksheet sheet = workbook.Worksheets[0];

foreach (OleObject ole in sheet.OleObjects)
    Console.WriteLine($"{ole.Label} from {ole.ObjectSourceFullName}, linked={ole.IsLink}");
```

## Delete an OLE object

`OleObjectCollection` has `RemoveAt(int)` and `Clear()`. Unlike pictures, there is no by-reference `Remove`.

```csharp
var workbook = new Workbook("has-ole.xlsx");
Worksheet sheet = workbook.Worksheets[0];

if (sheet.OleObjects.Count > 0)
    sheet.OleObjects.RemoveAt(0);
sheet.OleObjects.Clear();
workbook.Save("out.xlsx", SaveFormat.Xlsx);
```

## Related
- images-and-shapes.md - pictures share the same geometry members (Width/Height/Top/Left/Placement).
- conversion-and-rendering.md - OLE objects render into PDF/image output.
