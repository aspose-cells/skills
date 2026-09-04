---
name: insert-image
description: Insert raster, vector, and SVG pictures into Aspose.Cells for Java worksheets — anchored to a cell, positioned absolutely, or used as the source for a picture hyperlink. Use whenever the user wants to put a picture / logo / signature into a worksheet.
applies_to: Java
keywords: [image, picture, insert, logo, picture collection, SVG, webp]
---

# Insert images

## Imports

```java
import com.aspose.cells.*;
```

## Quick example — picture anchored to a cell

```java
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().get(0);
int idx = ws.getPictures().add(5, 0, "logo.png");  // row, column
Picture pic = ws.getPictures().get(idx);

pic.setTitle("Our Logo");
pic.setWidthScale(80);                            // 80 %
pic.setHeightScale(80);
```

## Positioned absolutely (pixel coordinates)

```java
// PictureCollection.add takes (upperLeftRow, upperLeftColumn, fileName)
// for the basic case; for pixel-precise placement, use addPicture.
int idx = ws.getPictures().add(0, 0, "signature.png");
Picture pic1 = ws.getPictures().get(idx);
pic1.setLeft(10);
pic1.setTop(10);
pic1.setWidth(200);
pic1.setHeight(60);
```

The `x` and `y` arguments are pixels relative to the top-left cell.

## Image format support

- PNG, JPEG, BMP, GIF, TIFF, EMF, WMF — out of the box
- **SVG** — convert to EMF first or use a vector wrapper
- **WebP** — requires the TwelveMonkeys ImageIO extension on the
  classpath. Add to `maven-gradle-dependencies.md` style block:
  `com.twelvemonkeys.imageio:imageio-webp`

```java
// WebP needs TwelveMonkeys ImageIO on the classpath; nothing to call at runtime.
```

## Picture hyperlink

```java
ws.getHyperlinks().add(5, 0, 5, 0, "https://aspose.com");   // (startRow, startCol, endRow, endCol, url)
```

## Pitfalls

- **Picture anchors** — adding a picture at row 5, column 0 places
  it on top of the cell; the upper-left corner of the picture aligns
  with the cell's top-left.
- **`addPicture(row, col, file)`** with a relative file path resolves
  against the JVM working directory, **not** the classpath — pass an
  absolute path or stream-based overloads.
- **Large images** balloon XLSX size — configure
  `opts.setQuality(quality)` after JPEG conversion if needed (0–100, lower = smaller file).
- **SVG in XLSX** — when Excel can't render SVG it falls back to an
  icon. Convert to PNG / EMF if preview safety matters.

## Related

- [shapes-textboxes.md](shapes-textboxes.md)
- [../_shared/fonts-linux-docker.md](../_shared/fonts-linux-docker.md)
