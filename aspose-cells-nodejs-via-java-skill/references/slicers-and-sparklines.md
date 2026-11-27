# Slicers and sparklines

## Slicers

A slicer belongs to a worksheet and filters a pivot table or ListObject table; it is not an independent filter over an arbitrary cell range. Create the source table/pivot first, select a valid field, and use the `SlicerCollection.add()` overload exposed by the installed release.

When editing an existing workbook, inspect slicers through the worksheet collection:

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook("slicers.xlsx");
const slicers = workbook.getWorksheets().get(0).getSlicers();
for (let i = 0; i < slicers.getCount(); i += 1) {
  console.log(slicers.get(i).getName());
}
```

For a pivot-backed slicer, filtering the cache does not guarantee that visible pivot output is already refreshed. Refresh the pivot cache and call `calculateData()` before reading or rendering updated results. Do not confuse a table slicer with worksheet AutoFilter or with a pivot report filter.

Slicer creation overloads and caption/style members are version-sensitive. Verify the required field argument and destination-cell overload against the installed Node API rather than translating a .NET property-based example mechanically.

## Sparklines

Sparkline groups define the data and destination area; individual sparklines map a data row/column to a destination cell.

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook();
const sheet = workbook.getWorksheets().get(0);
sheet.getCells().get("A2").putValue(10);
sheet.getCells().get("B2").putValue(15);
sheet.getCells().get("C2").putValue(12);

const area = new cells.CellArea();
area.setStartRow(1);
area.setEndRow(1);
area.setStartColumn(3);
area.setEndColumn(3);
const index = sheet.getSparklineGroups().add(
  cells.SparklineType.LINE,
  "A2:C2",
  false,
  area
);
const group = sheet.getSparklineGroups().get(index);
group.getSparklines().add("A2:C2", 1, 3);
group.setShowHighPoint(true);

const color = workbook.createCellsColor();
color.setColor(cells.Color.getBlue());
group.setSeriesColor(color);
workbook.save("sparklines.xlsx", cells.SaveFormat.XLSX);
```

Destination cell indices in `SparklineCollection.add()` are 0-based. Sparkline colors use `CellsColor`, created by the workbook, rather than assigning `Color` directly. Ensure the data range orientation matches the `isVertical` argument and that destination cells do not overlap the source data.
