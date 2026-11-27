# Formulas and calculation

## Recalculate before reading results

Setting a formula does not evaluate it. Read a computed value only after calculation.

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook();
const sheet = workbook.getWorksheets().get(0);
sheet.getCells().get("A1").putValue(10);
sheet.getCells().get("A2").putValue(20);
sheet.getCells().get("A3").setFormula("=SUM(A1:A2)");

workbook.calculateFormula();
console.log(sheet.getCells().get("A3").getValue());
workbook.save("calculated.xlsx", cells.SaveFormat.XLSX);
```

Use invariant Excel formula text: English function names, comma argument separators, and A1 references unless deliberately using R1C1. Do not build formulas with locale-specific decimal separators.

## Editing formulas

Re-run calculation after changing a formula or any precedent value. A previously cached result can otherwise look plausible but be stale. A newly added formula can return `null`, zero, or an old cached value until this happens.

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook("input.xlsx");
const sheet = workbook.getWorksheets().get(0);
sheet.getCells().get("C1").setFormula("=A1*B1");
sheet.getCells().get("C2").setR1C1Formula("=RC[-2]*RC[-1]");
workbook.calculateFormula();
```

## Defined names

Use workbook-level names through the worksheet collection.

```javascript
const cells = require("aspose.cells");

const workbook = new cells.Workbook();
const names = workbook.getWorksheets().getNames();
const index = names.add("TaxRate");
names.get(index).setRefersTo("=Sheet1!$B$1");
workbook.getWorksheets().get("Sheet1").getCells().get("B1").putValue(0.2);
workbook.getWorksheets().get("Sheet1").getCells().get("C1").setFormula("=A1*TaxRate");
workbook.calculateFormula();
```

Quote sheet names containing spaces or punctuation in both formulas and `refersTo` strings, for example `='Sales Data'!$A$1`.

## Calculation failures and advanced features

- Check `#REF!`, `#NAME?`, and `#VALUE!` results at their source rather than replacing them with zero.
- Circular references require an intentional iterative-calculation configuration; do not hide them with repeated recalculation.
- Large dependency chains are expensive. Batch edits, then calculate once rather than after every cell.
- Calculation options, custom functions, and array-formula overloads are version-sensitive. Confirm the installed Node API and bridge support before using Java interfaces such as custom-function implementations.
- Aspose.Cells calculates formulas in memory. Rendering or saving does not make every external consumer refresh them; retain cached values only after successful calculation.
