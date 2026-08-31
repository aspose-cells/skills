# Images, Shapes, and Comments

Read this when inserting or positioning pictures and logos, sizing an image to a cell, adding shapes or text boxes, or attaching cell comments and notes.

Every drawing collection hangs off `Worksheet`: `sheet.pictures`, `sheet.shapes`, `sheet.text_boxes`, `sheet.comments`. All row and column arguments are 0-based (row index 5 = Excel row 6, column index 2 = column C). The three positioning models are anchor cells, pixel offsets, and percent scale.

## Pitfalls

### Picture lands in the wrong cell, or shifts and stretches when rows change

Symptom: the logo appears one cell off from where you wanted it, or it jumps and resizes after you insert, delete, or resize rows and columns.

Cause: `Pictures.add(row, column, ...)` anchors the picture's top-left corner at a 0-based cell, so passing an Excel 1-based number lands it one row and column too far. Separately, `Picture.placement` decides how the picture reacts to cell edits; by default it moves and sizes with its anchor cells.

Fix: pass 0-based indices, then set `placement` explicitly. `PlacementType` has three values: `FREE_FLOATING` (never move or size with cells), `MOVE` (move but keep its size), `MOVE_AND_SIZE` (move and resize with the cells).

```python
import aspose.cells as gc
import aspose.cells.drawing as dw

workbook = gc.Workbook()
sheet = workbook.worksheets[0]

# Anchor top-left at row index 5 (Excel row 6), column index 1 (column B)
index = sheet.pictures.add(5, 1, "logo.png")
picture = sheet.pictures[index]

# Pin it so editing rows never moves or stretches it
picture.placement = dw.PlacementType.FREE_FLOATING

workbook.save("out.xlsx", gc.SaveFormat.XLSX)
```

### Image aspect looks squashed when forced into a cell

Symptom: the picture is visibly stretched or compressed after you fit it to a cell.

Cause: setting `width` and `height` independently, or using the four-corner anchor overload, forces the image into the target rectangle and ignores the source aspect ratio.

Fix: to keep aspect, scale by a single factor with `width_scale` equal to `height_scale` (percent of the original), or set only one of `width`/`height`. When exact cell coverage matters more than aspect, use the four-anchor `add` overload with `MOVE_AND_SIZE`, or embed the image in the cell (see below) so it auto-scales.

```python
import io
import aspose.cells as gc
import aspose.cells.drawing as dw

workbook = gc.Workbook()
sheet = workbook.worksheets[0]

with open("logo.png", "rb") as f:
    data = f.read()
# 5-arg overload: (upper_row, upper_col, lower_row, lower_col, source)
# source is a file-path string OR a Stream (io.BytesIO) - NOT raw bytes
i = sheet.pictures.add(5, 2, 6, 3, io.BytesIO(data))   # covers exactly cell C6
sheet.pictures[i].placement = dw.PlacementType.MOVE_AND_SIZE
workbook.save("out.xlsx", gc.SaveFormat.XLSX)
```

## Insert and position pictures

`PictureCollection.add` has overloads taking a file path **string** or a `Stream` (e.g. `io.BytesIO`), an optional lower-right anchor, and an optional `width_scale`/`height_scale` pair:

- `add(row, column, "logo.png")` and `add(row, column, stream)` - anchor the top-left corner only.
- `add(upper_left_row, upper_left_column, lower_right_row, lower_right_column, source)` - anchor both corners so the picture spans that cell range.
- `add(row, column, source, width_scale, height_scale)` - anchor top-left and scale in percent of the original.

Position an already-added picture in pixels. `left`/`top` are the offset inside the anchor cell; `x`/`y` are the offset from the worksheet's top-left corner. `width_scale`/`height_scale` resize in percent of the original image.

```python
import aspose.cells as gc

workbook = gc.Workbook()
sheet = workbook.worksheets[0]
i = sheet.pictures.add(5, 5, "logo.png")
picture = sheet.pictures[i]

picture.left = 60   # pixels inside the anchor cell
picture.top = 10

picture.width_scale = 50    # 50% of original, keeps aspect ratio
picture.height_scale = 50
```

## Fit an image inside one cell

Two approaches. A floating picture covering one cell: use the four-anchor overload with `lower_right_row = r + 1` and `lower_right_column = c + 1`, plus `placement = MOVE_AND_SIZE` (shown above). Or embed the image as real cell content via `Cell.embedded_image` (a `bytes` value), which auto-scales to the cell and travels with it when you sort, filter, or move rows.

```python
import aspose.cells as gc

workbook = gc.Workbook()
sheet = workbook.worksheets[0]

with open("logo.png", "rb") as f:
    data = f.read()
cell = sheet.cells.get("C6")
cell.embedded_image = data

# Enlarge the cell so the in-cell image is visible
sheet.cells.set_column_width(2, 30.0)
sheet.cells.set_row_height(5, 100.0)

workbook.save("out.xlsx", gc.SaveFormat.XLSX)
```

## Access and remove pictures

Index with `sheet.pictures[i]`; get the size with `sheet.pictures.length` (or `len(sheet.pictures)`). The collection exposes `remove(picture)` (pass the `Picture` object, not an index) and `clear()`.

```python
import aspose.cells as gc

workbook = gc.Workbook()
sheet = workbook.worksheets[0]
sheet.pictures.add(0, 0, "logo.png")

if sheet.pictures.length > 0:
    picture = sheet.pictures[0]   # access by index
    picture.placement = gc.drawing.PlacementType.MOVE
sheet.pictures.remove(sheet.pictures[0])   # remove one (pass the Picture object)
sheet.pictures.clear()         # remove all
```

## Text boxes and shapes

`sheet.text_boxes.add(top_row, left_column, height, width)` returns an index; set `TextBox.text` on the item. `sheet.shapes` builds the rest: `add_rectangle`, `add_line`, and `add_auto_shape(AutoShapeType, ...)` all take `(upper_left_row, top, upper_left_column, left, height, width)`, where `top`/`left` are pixel offsets inside the anchor cell and `height`/`width` are pixel sizes.

```python
import aspose.cells as gc
import aspose.cells.drawing as dw

workbook = gc.Workbook()
sheet = workbook.worksheets[0]

# TextBox anchored at (row 1, col 1), 80px tall, 200px wide
t = sheet.text_boxes.add(1, 1, 80, 200)
sheet.text_boxes[t].text = "Quarterly report"

# Shapes: (upper_left_row, top, upper_left_column, left, height, width)
sheet.shapes.add_rectangle(6, 0, 1, 0, 60, 180)
sheet.shapes.add_line(10, 0, 1, 0, 0, 180)
sheet.shapes.add_auto_shape(dw.AutoShapeType.CUBE, 12, 0, 1, 0, 80, 80)

workbook.save("out.xlsx", gc.SaveFormat.XLSX)
```

## Cell comments and notes

`sheet.comments.add("A1")` or `sheet.comments.add(row, column)` returns the new comment's index. Set `note` for the text, plus optional `author`, `width`/`height` (pixels), and `font`.

```python
import aspose.cells as gc

workbook = gc.Workbook()
sheet = workbook.worksheets[0]

c = sheet.comments.add("A1")
comment = sheet.comments[c]
comment.note = "Check this figure"
comment.author = "Finance"
comment.width = 200   # pixels
comment.height = 100

workbook.save("out.xlsx", gc.SaveFormat.XLSX)
```

### Threaded comments (Excel 365 conversations)

Threaded comments need a `ThreadedCommentAuthor`. Add the author to `workbook.worksheets.threaded_comment_authors`, then call `add_threaded_comment(cell_name, text, author)`. Read them back with `sheet.comments.get_threaded_comments("A1")`, which returns a collection you can iterate.

```python
import aspose.cells as gc

workbook = gc.Workbook()
sheet = workbook.worksheets[0]

a = workbook.worksheets.threaded_comment_authors.add("Jane Doe", "", "")
author = workbook.worksheets.threaded_comment_authors[a]
sheet.comments.add_threaded_comment("A1", "Please review Q3", author)

workbook.save("out.xlsx", gc.SaveFormat.XLSX)
```

## Not covered here

- Images in page headers and footers, and print setup -> conversion-and-rendering.md
- Rendering a chart to an image (PNG or SVG) -> charts.md

## Related
- conversion-and-rendering.md - header/footer images and print setup.
- charts.md - chart-to-image rendering.
- styles-and-formatting.md - formatting the cells around pictures.
- worksheets-rows-columns.md - row/column sizes that anchor pictures.
