---
name: set-get-formula
description: Set and read cell formulas in Aspose.Cells for Java — including built-in functions, R1C1 references, array formulas, and shared/shared formula management. Always call calculateFormula() afterwards or values will be blank.
applies_to: Java
keywords: [formula, setFormula, getFormula, calculate, function, builtin]
---

# Set & get formulas

## Imports

```java
import com.aspose.cells.*;
```

## Quick example

```java
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().get(0);

ws.getCells().get("A1").setFormula("=B1+C1");
ws.getCells().get("B1").setValue(10);
ws.getCells().get("C1").setValue(20);

wb.calculateFormula();                          // mandatory
System.out.println(ws.getCells().get("A1").getValue());  // 30
```

Without `calculateFormula()`, `getValue()` is `null` and the saved
file would show "0" or "#NAME?".

## Read formula text vs. result

```java
Worksheet ws = wb.getWorksheets().get(0);
Cell cell = ws.getCells().get("A1");
String formula = cell.getFormula();             // text starting with "="
Object value   = cell.getValue();               // last computed/calculated value
double num     = cell.getDoubleValue();         // numeric result if available
```

## Built-in functions

Aspose.Cells ships with Excel-compatible implementations of nearly all
functions. Use them by name inside the `=` string.

```java
ws.getCells().get("D1").setFormula("=SUM(B1:C10)");
ws.getCells().get("D2").setFormula("=VLOOKUP(A2, Sheet2!A:F, 5, FALSE)");
ws.getCells().get("D3").setFormula("=IF(B3>0, \"OK\", \"OVER\")");
```

## Add-in / user-defined functions

For `.xlam`-style add-ins:

```java
wb.getWorksheets().registerAddInFunction("MyFunc.xlam", "MY_FUNC", true);
ws.getCells().get("Z1").setFormula("=MY_FUNC(A1)");
```

For custom Java engines see `../formulas/calculate-formulas.md`.

## Set a formula and write back the cached value

```java
Worksheet ws = wb.getWorksheets().get(0);
Cell cell = ws.getCells().get("A1");
cell.setFormula("=A1*2");
cell.putValue(cell.getDoubleValue());           // copy the cached numeric value
```

Use only when the user has hand-computed values; otherwise let
`calculateFormula` populate them.

## Pitfalls

- **Always `calculateFormula()`** before reading `getValue()` or
  before saving, otherwise outputs are blank / `#NAME?`.
- **Formula strings start with `=`** — `setFormula("A1+B1")` is
  treated as plain text.
- **Locale conflicts** — formulas use `,` parameter separator
  (matching Excel); not `;`, regardless of region.
- **Volatile formulas** (`=NOW()`, `=RAND()`) are recomputed every
  open; avoid in test data.
- **Excel-2021+ functions** (`UNIQUE`, `FILTER`, `LET`, `SORT`, …) —
  Aspose.Cells persists the formula string and lets the user's Excel
  render them at open. There is no separate "new calculation engine"
  flag to enable; if the renderer does not support them, add
  `ForceFullCalculation` or a manual `calculateFormula()` call after
  edits so any per-row spill is re-evaluated.

## Related

- [calculate-formulas.md](calculate-formulas.md)
- [array-r1c1-formula.md](array-r1c1-formula.md)
- [named-ranges.md](named-ranges.md)
