---
name: shapes-textboxes
description: Add Excel shapes, text boxes, OLE objects, and WordArt to Aspose.Cells for Java worksheets. Use whenever the user asks to put a shape, an embedded object, or a watermark-text box on a sheet.
applies_to: Java
keywords: [shape, text box, OLE, WordArt, MsoDrawingType, drawing]
---

# Shapes, text boxes & OLE

## Imports

```java
import com.aspose.cells.*;
```

## Quick example — text box

```java
import static com.aspose.cells.Color.getRed;

Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().get(0);

int idx = ws.getTextBoxes().add(2, 2, 200, 80);   // row, col, w, h
TextBox tb = ws.getTextBoxes().get(idx);
tb.setText("Confidential");
tb.getFont().setSize(18);
tb.getFont().setBold(true);
tb.getFont().setColor(getRed());

tb.setPlacement(PlacementType.FREE_FLOATING);     // move freely, not anchored
```

## Generic shape (line, rectangle, oval, …)

```java
Shape rect = ws.getShapes().addShape(MsoDrawingType.RECTANGLE, 0, 0, 0, 0, 200, 50);
```

## OLE object

```java
// row, column, height, width, imageData
byte[] doc = new byte[0];                          // replace with real file bytes
int idx2 = ws.getOleObjects().add(8, 0, 200, 80, doc);
OleObject ole = ws.getOleObjects().get(idx2);
ole.setDisplayAsIcon(true);                       // Word icon, double-click to open
```

## WordArt

```java
int wordIdx = ws.getTextBoxes().add(0, 0, 100, 50);
TextBox wordArt = ws.getTextBoxes().get(wordIdx);
wordArt.setText("Total Revenue");
wordArt.getFont().setSize(24);
```

## Pitfalls

- **Index `add(...)` → `get(...)`** because `add` returns the *index*,
  not the object. Mixing them up is the #1 shape bug.
- **`PlacementType.FREE_FLOATING` vs `MOVE_AND_SIZE`** — `MOVE_AND_SIZE`
  (default) anchors to the cell, resize-with-cell; `FREE_FLOATING` is freely
  floating and survives row / column insertion.
- **Group shapes** — call `ws.getShapes().group(groupShapesArray)`
  to combine multiple shapes into one.
- **OLE objects** are embedded *as files* by Excel — calling save
  again with a different cell link keeps the original bytes; this
  is how VBA / formulas can store binary data.

## Related

- [insert-image.md](insert-image.md)
