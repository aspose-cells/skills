# Smart markers reporting

Smart markers fill a workbook template through `WorkbookDesigner`. Markers use `&=` syntax and are stored as template cell text.

## Scalar marker

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook();
workbook.getWorksheets().get(0).getCells().get("A1").putValue("&=Customer");

const designer = new cells.WorkbookDesigner();
designer.setWorkbook(workbook);
designer.setDataSource("Customer", "Ada");
designer.process();
workbook.save("report.xlsx", cells.SaveFormat.XLSX);
```

The data-source name must match the marker exactly. Set all sources before calling `process()` and save the same workbook held by the designer.

## Repeating data

Markers such as `&=Orders.Product` require a compatible object/collection data source and a template laid out for repeated rows. In the Java product, common sources include JavaBeans and supported Java collections. Node via Java does not make a plain JavaScript object a JavaBean or a JavaScript array a Java array.

Before generating a repeating-report implementation:

1. Identify the exact `setDataSource()` overload exposed by the installed release.
2. Construct the documented Java-compatible source through the bridge, or perform explicit per-cell/template-row writes instead.
3. Test field/property resolution, empty collections, null values, and row insertion behavior.
4. Recalculate formulas after processing if formulas depend on inserted data.

Do not pass `JSON.parse()` output or a normal JavaScript array and assume designer reflection will work. The scalar example above must not be generalized into an unsupported collection example.

## Troubleshooting

- Marker remains unchanged: source name/property mismatch, unsupported source shape, or `process()` was omitted.
- Rows overwrite each other: template marker layout and grouping are wrong for a repeating source.
- Formulas show old results: calculate after processing.
- Styles disappear: place the required style on the template row and verify the designer's processing options for the installed release.
- Extra blank rows: check marker cells, source emptiness, and template range boundaries before adding cleanup logic.
