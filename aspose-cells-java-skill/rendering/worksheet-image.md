---
name: worksheet-image
description: Render an Aspose.Cells for Java worksheet to PNG, JPEG, BMP, TIFF, GIF — ImageOrPrintOptions for resolution, transparency, smoothing, area, page index. Use when the user wants any image export of one or many sheets.
applies_to: Java
keywords: [image, png, jpeg, tiff, bmp, gif, ImageOrPrintOptions, SheetRender]
---

# Worksheet → image

## Imports

```java
import com.aspose.cells.*;
```

## Quick example — one PNG per sheet

```java
Workbook wb = new Workbook("report.xlsx");

ImageOrPrintOptions imgOpts = new ImageOrPrintOptions();
imgOpts.setImageType(ImageType.PNG);
imgOpts.setOnePagePerSheet(true);
imgOpts.setHorizontalResolution(200);
imgOpts.setVerticalResolution(200);

for (int s = 0; s < wb.getWorksheets().getCount(); s++) {
    Worksheet ws = wb.getWorksheets().get(s);
    SheetRender renderer = new SheetRender(ws, imgOpts);
    for (int p = 0; p < renderer.getPageCount(); p++) {
        renderer.toImage(p, ws.getName().replace(' ', '_') + "-" + p + ".png");
    }
}
```

## Output types

| Image type     | Enum value                |
| -------------- | ------------------------- |
| PNG            | `ImageType.PNG`           |
| JPEG           | `ImageType.JPEG`          |
| BMP            | `ImageType.BMP`           |
| GIF            | `ImageType.GIF`           |
| TIFF           | `ImageType.TIFF`          |
| EMF / SVG / …  | `ImageType.EMF` etc.      |

For TIFF quality use `TiffCompression` enum and
`imgOpts.setTiffCompression(...)`.

## Specific area only

```java
ImageOrPrintOptions imgOpts = new ImageOrPrintOptions();
imgOpts.setOnlyArea(true);                                  // restrict to print area
ws.getPageSetup().setPrintArea("A1:D20");                   // rows 0..19, cols 0..3
```

## Transparency

```java
ImageOrPrintOptions imgOpts = new ImageOrPrintOptions();
imgOpts.setImageType(ImageType.PNG);
imgOpts.setTransparent(true);                     // keep sheet transparent areas
```

For finer control, post-process the rendered PNG bytes and apply your
own alpha channel.

## Pitfalls

- **ImageType.TIFF on JDK < 9** requires JAI / JAI Image I/O Tools
  on the classpath. Otherwise a `NoClassDefFoundError` is thrown.
  Prefer PNG/JPEG unless TIFF is mandatory.
- **`setOnePagePerSheet(true)` enlarges tiny sheets**; clear the
  flag if you want the natural-sized result.
- **CSS-rendering is not involved** — fonts that are missing from the
  container are missing. See `../_shared/fonts-linux-docker.md`.
- **Multi-page sheets** require looping over `renderer.getPageCount()`
  and calling `toImage(page, stream)` per page.
- **Setting only `setHorizontalResolution`** without vertical uses
  the same value for both. Set both for predictable DPI.

## Related

- [convert-to-pdf.md](convert-to-pdf.md)
- [convert-html.md](convert-html.md)
- [../_shared/fonts-linux-docker.md](../_shared/fonts-linux-docker.md)
