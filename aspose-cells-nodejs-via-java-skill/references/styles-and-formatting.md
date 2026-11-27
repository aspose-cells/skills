# Styles and formatting

## Cell styles are copied

`getStyle()` returns a style copy. Mutate it and call `setStyle()`; otherwise the change does not stick.

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook();
const cell = workbook.getWorksheets().get(0).getCells().get("A1");
cell.putValue("Total");

const style = cell.getStyle();
style.getFont().setBold(true);
style.getFont().setColor(cells.Color.getDarkBlue());
style.setCustom("#,##0.00");
cell.setStyle(style);
```

Built-in numeric formats can be selected with `style.setNumber(formatId)`. Use `setCustom()` for an explicit Excel format such as `yyyy-mm-dd`, `0.00%`, or `#,##0.00`. A stored value and its display format are separate concerns.

## Reuse a style for ranges

Do not create an independent style for every cell in a large loop. Excel limits the number of unique cell formats. Create once and apply with a `StyleFlag` that identifies the attributes to copy.

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook();
const sheet = workbook.getWorksheets().get(0);
const style = workbook.createStyle();
style.getFont().setBold(true);
style.setCustom("#,##0.00");
style.setHorizontalAlignment(cells.TextAlignmentType.RIGHT);

const flags = new cells.StyleFlag();
flags.setAll(true);
sheet.getCells().createRange("A1:C20").applyStyle(style, flags);
```

Use narrower flags instead of `setAll(true)` when the operation must preserve existing borders, fills, or number formats.

## Borders and fills

Style components are mutable objects reached through getters. Configure them before writing the style back to a cell or applying it to a range. Prefer named/static Aspose `Color` values known to the installed release over constructing colors with an untested argument order.

## Conditional formatting

Conditional formatting belongs to a worksheet and applies to one or more `CellArea` instances.

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook();
const sheet = workbook.getWorksheets().get(0);
sheet.getCells().get("A1").putValue(75);

const index = sheet.getConditionalFormattings().add();
const conditions = sheet.getConditionalFormattings().get(index);
const area = new cells.CellArea();
area.setStartRow(0);
area.setEndRow(19);
area.setStartColumn(0);
area.setEndColumn(0);
conditions.addArea(area);

const conditionIndex = conditions.addCondition(
  cells.FormatConditionType.CELL_VALUE,
  cells.OperatorType.GREATER_THAN,
  "50",
  ""
);
const condition = conditions.get(conditionIndex);
condition.getStyle().getFont().setBold(true);
```

Verify operator-specific formula arguments against the installed API; operators that need one value must not be given a made-up second expression. Conditional formatting is evaluated by Excel/compatible consumers, and it does not replace explicit formula recalculation when another API reads cached formula results.
