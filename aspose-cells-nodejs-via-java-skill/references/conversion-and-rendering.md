# Conversion and rendering

## Save spreadsheet formats

Use an explicit save format even when the filename has an extension.

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook();
workbook.getWorksheets().get(0).getCells().get("A1").putValue("Hello");
workbook.save("output.xlsx", cells.SaveFormat.XLSX);
```

A macro-enabled workbook must be saved to a macro-enabled format such as XLSM. Conversion to XLSX removes VBA content.

## PDF

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook("input.xlsx");
const options = new cells.PdfSaveOptions();
options.setOnePagePerSheet(true);
workbook.save("output.pdf", options);
```

`onePagePerSheet` is a deliberate scaling choice and can make large sheets unreadable. For ordinary pagination, configure each worksheet instead:

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook("input.xlsx");
const setup = workbook.getWorksheets().get(0).getPageSetup();
setup.setOrientation(cells.PageOrientationType.LANDSCAPE);
setup.setPaperSize(cells.PaperSizeType.A4);
setup.setFitToPagesWide(1);
setup.setFitToPagesTall(0);
workbook.save("paged.pdf", new cells.PdfSaveOptions());
```

Recalculate formulas and charts before rendering. Install the workbook's fonts and compare page count and text clipping in the production environment.

## HTML

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook("input.xlsx");
const options = new cells.HtmlSaveOptions();
options.setExportActiveWorksheetOnly(true);
workbook.save("output.html", options);
```

HTML export can create related resources. Return the complete documented output rather than only the main file when images or other assets are emitted.

## Worksheet images

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook("input.xlsx");
const sheet = workbook.getWorksheets().get(0);
const options = new cells.ImageSaveOptions(cells.SaveFormat.PNG);
const renderer = new cells.SheetRender(sheet, options);
renderer.toImage(0, "sheet-1.png");
```

Check `renderer` page count for multi-page sheets. Image conversion is paginated rendering, not a screenshot of Excel.

## CSV and TSV

Text export normally targets the active worksheet unless configured otherwise. Select the active sheet deliberately or use `TxtSaveOptions.setExportAllSheets(true)` where supported. Encoding, delimiter, quote handling, dates, formulas, and multiple sheets require explicit decisions; XLSX-only features cannot survive CSV.

## Java-compatible streams

Node via Java stream operations differ from ordinary path-based saves and from Node stream conventions. The documented open operation is asynchronous and invokes its callback with `(workbook, error)`:

```javascript
const cells = require("aspose.cells");

function openFromJavaStream(inputStream) {
  cells.Workbook.createWorkbookFromStream(inputStream, (workbook, error) => {
    if (error) {
      console.error(error);
      return;
    }
    workbook.save("from-stream.xlsx", cells.SaveFormat.XLSX);
  });
}
```

Static stream saving uses the workbook as the first argument:

```javascript
const cells = require("aspose.cells");

function saveToJavaStream(workbook, outputStream) {
  cells.Workbook.saveToStream(workbook, outputStream, cells.SaveFormat.XLSX);
}
```

Only pass stream implementations accepted by the installed API. Do not substitute `fs.createReadStream()`, `fs.createWriteStream()`, or a web response without an explicit adapter documented for this binding. Propagate stream completion and errors through the surrounding application rather than returning before the callback completes.

## Printing

Node via Java does not support Aspose's Java printing APIs. Render to PDF or images and hand the result to an operating-system or service-side printing component.
