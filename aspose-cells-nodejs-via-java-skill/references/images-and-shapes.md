# Images and shapes

## Insert a picture

Picture coordinates are 0-based cell coordinates. Width and height are in pixels.

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook();
const sheet = workbook.getWorksheets().get(0);
const index = sheet.getPictures().add(1, 1, "logo.png");
const picture = sheet.getPictures().get(index);
picture.setWidth(160);
picture.setHeight(80);
workbook.save("picture.xlsx", cells.SaveFormat.XLSX);
```

Setting both dimensions can distort the image. Preserve the source aspect ratio when only one dimension should drive resizing. Choose placement behavior deliberately when rows/columns can move or resize; do not assume all consumers interpret anchor behavior identically.

For a service, resolve the source path explicitly and verify that the file exists before adding it. If an overload takes image bytes or a stream, verify that exact type with the Node via Java bridge; do not assume a JavaScript `Buffer` or Node stream is accepted.

## Add a text box

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook();
const sheet = workbook.getWorksheets().get(0);
const textBox = sheet.getShapes().addTextBox(1, 0, 1, 0, 80, 220);
textBox.getTextBody().setText("Prepared by Aspose.Cells");
workbook.save("textbox.xlsx", cells.SaveFormat.XLSX);
```

Shape arguments include row/column anchors, pixel offsets, height, and width. Consult the installed overload signature before switching to another `addShape` form.

## Comments

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook();
const sheet = workbook.getWorksheets().get(0);
const index = sheet.getComments().add("A1");
sheet.getComments().get(index).setNote("Reviewed value");
workbook.save("comments.xlsx", cells.SaveFormat.XLSX);
```

Comments are worksheet objects, not cell values. Modern threaded comments and legacy notes are not interchangeable; verify support in the target version and output format.

## Rendering reminders

Fonts affect text-box measurement and image/PDF layout. Install the fonts used by shapes and charts on the production host. After modifying chart-backed shapes, follow the chart calculation guidance in [charts](charts.md).
