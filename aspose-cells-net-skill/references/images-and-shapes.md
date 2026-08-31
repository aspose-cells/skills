# Images, Shapes, and Comments

Read this when inserting or positioning pictures and logos, sizing an image to a cell, adding shapes or text boxes, or attaching cell comments and notes.

Every drawing collection hangs off `Worksheet`: `sheet.Pictures`, `sheet.Shapes`, `sheet.TextBoxes`, `sheet.Comments`. All row and column arguments are 0-based (row index 5 = Excel row 6, column index 2 = column C). The three positioning models are anchor cells, pixel offsets, and percent scale.

## Pitfalls

### Picture lands in the wrong cell, or shifts and stretches when rows change

Symptom: the logo appears one cell off from where you wanted it, or it jumps and resizes after you insert, delete, or resize rows and columns.

Cause: `Pictures.Add(row, column, ...)` anchors the picture's top-left corner at a 0-based cell, so passing an Excel 1-based number lands it one row and column too far. Separately, `Picture.Placement` decides how the picture reacts to cell edits; by default it moves and sizes with its anchor cells.

Fix: pass 0-based indices, then set `Placement` explicitly. `PlacementType` has three values: `FreeFloating` (never move or size with cells), `Move` (move but keep its size), `MoveAndSize` (move and resize with the cells).

```csharp
var workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];

// Anchor top-left at row index 5 (Excel row 6), column index 1 (column B)
int index = sheet.Pictures.Add(5, 1, "logo.png");
var picture = sheet.Pictures[index];

// Pin it so editing rows never moves or stretches it
picture.Placement = PlacementType.FreeFloating;

workbook.Save("out.xlsx", SaveFormat.Xlsx);
```

### Image aspect looks squashed when forced into a cell

Symptom: the picture is visibly stretched or compressed after you fit it to a cell.

Cause: setting `Width` and `Height` independently, or using the four-corner anchor overload, forces the image into the target rectangle and ignores the source aspect ratio.

Fix: to keep aspect, scale by a single factor with `WidthScale` equal to `HeightScale` (percent of the original), or set only one of `Width`/`Height`. When exact cell coverage matters more than aspect, use the four-anchor `Add` overload with `MoveAndSize`, or embed the image in the cell (see below) so it auto-scales.

```csharp
var workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];

byte[] bytes = File.ReadAllBytes("logo.png");
using (var ms = new MemoryStream(bytes))
{
    // Cover exactly cell C6: upper-left (5,2), lower-right (6,3)
    int i = sheet.Pictures.Add(5, 2, 6, 3, ms);
    sheet.Pictures[i].Placement = PlacementType.MoveAndSize;
}
workbook.Save("out.xlsx", SaveFormat.Xlsx);
```

## Insert and position pictures

`PictureCollection.Add` has overloads for a file path or a `Stream`, an optional lower-right anchor, and an optional `widthScale`/`heightScale` pair:

- `Add(row, column, "logo.png")` and `Add(row, column, stream)` - anchor the top-left corner only.
- `Add(upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn, source)` - anchor both corners so the picture spans that cell range.
- `Add(row, column, source, widthScale, heightScale)` - anchor top-left and scale in percent of the original.

Position an already-added picture in pixels. `Left`/`Top` are the offset inside the anchor cell; `X`/`Y` are the offset from the worksheet's top-left corner. `WidthScale`/`HeightScale` resize in percent of the original image.

```csharp
var workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];
int i = sheet.Pictures.Add(5, 5, "logo.png");
var picture = sheet.Pictures[i];

picture.Left = 60;   // pixels inside the anchor cell
picture.Top = 10;

picture.WidthScale = 50;    // 50% of original, keeps aspect ratio
picture.HeightScale = 50;
```

## Fit an image inside one cell

Two approaches. A floating picture covering one cell: use the four-anchor overload with `LowerRightRow = r + 1` and `LowerRightColumn = c + 1`, plus `Placement = PlacementType.MoveAndSize` (shown above). Or embed the image as real cell content via `Cell.EmbeddedImage` (a `byte[]`), which auto-scales to the cell and travels with it when you sort, filter, or move rows.

```csharp
var workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];

Cell cell = sheet.Cells["C6"];
cell.EmbeddedImage = File.ReadAllBytes("logo.png");

// Enlarge the cell so the in-cell image is visible
sheet.Cells.SetColumnWidth(2, 30);
sheet.Cells.SetRowHeight(5, 100);

workbook.Save("out.xlsx", SaveFormat.Xlsx);
```

## Access and remove pictures

Index with `sheet.Pictures[i]`, count with `sheet.Pictures.Count`. The collection exposes `RemoveAt(int)` and `Clear()`; there is no `Remove(picture)`.

```csharp
var workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];
sheet.Pictures.Add(0, 0, "logo.png");

if (sheet.Pictures.Count > 0)
{
    var picture = sheet.Pictures[0];   // access by index
    picture.Placement = PlacementType.Move;
}
sheet.Pictures.RemoveAt(0);   // remove one by index
sheet.Pictures.Clear();       // remove all
```

## Text boxes and shapes

`sheet.TextBoxes.Add(topRow, leftColumn, height, width)` returns an index; set `TextBox.Text` on the item. `sheet.Shapes` builds the rest: `AddRectangle`, `AddLine`, and `AddAutoShape(AutoShapeType, ...)` all take `(upperLeftRow, top, upperLeftColumn, left, height, width)`, where `top`/`left` are pixel offsets inside the anchor cell and `height`/`width` are pixel sizes.

```csharp
var workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];

// TextBox anchored at (row 1, col 1), 80px tall, 200px wide
int t = sheet.TextBoxes.Add(1, 1, 80, 200);
sheet.TextBoxes[t].Text = "Quarterly report";

// Shapes: (upperLeftRow, top, upperLeftColumn, left, height, width)
sheet.Shapes.AddRectangle(6, 0, 1, 0, 60, 180);
sheet.Shapes.AddLine(10, 0, 1, 0, 0, 180);
sheet.Shapes.AddAutoShape(AutoShapeType.Cube, 12, 0, 1, 0, 80, 80);

workbook.Save("out.xlsx", SaveFormat.Xlsx);
```

## Cell comments and notes

`sheet.Comments.Add("A1")` or `sheet.Comments.Add(row, column)` returns the new comment's index. Set `Note` for the text, plus optional `Author`, `Width`/`Height` (pixels), and `Font`.

```csharp
var workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];

int c = sheet.Comments.Add("A1");
Comment comment = sheet.Comments[c];
comment.Note = "Check this figure";
comment.Author = "Finance";
comment.Width = 200;   // pixels
comment.Height = 100;

workbook.Save("out.xlsx", SaveFormat.Xlsx);
```

### Threaded comments (Excel 365 conversations)

Threaded comments need a `ThreadedCommentAuthor`. Add the author to `workbook.Worksheets.ThreadedCommentAuthors`, then call `AddThreadedComment(cellName, text, author)`. Read them back with `sheet.Comments.GetThreadedComments("A1")`, which returns a collection you can iterate.

```csharp
var workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];

int a = workbook.Worksheets.ThreadedCommentAuthors.Add("Jane Doe", "", "");
var author = workbook.Worksheets.ThreadedCommentAuthors[a];
sheet.Comments.AddThreadedComment("A1", "Please review Q3", author);

workbook.Save("out.xlsx", SaveFormat.Xlsx);
```

## Not covered here

- Images in page headers and footers, and print setup -> conversion-and-rendering.md
- Rendering a chart to an image (PNG or SVG) -> charts.md

## Related
- conversion-and-rendering.md - header/footer images and print setup.
- charts.md - chart-to-image rendering.
- styles-and-formatting.md - formatting the cells around pictures.
- worksheets-rows-columns.md - row/column sizes that anchor pictures.
