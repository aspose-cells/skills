---
name: array-r1c1-formula
description: Set array formulas and R1C1-reference formulas in Aspose.Cells for Java — including DATA_TABLE / LET / FILTER (Excel 2021+) behaviour and how to compute spill ranges. Use when the user wants a CSE-style formula or A1 style replaced with R1C1.
applies_to: Java
keywords: [array formula, CSE, R1C1, dynamic array, spill]
---

# Array & R1C1 formulas

## Imports

```java
import com.aspose.cells.*;
```

## Array formula

```java
Cell topLeft = ws.getCells().get("G2");
topLeft.setArrayFormula("=SUM(A2:D2*A2:D2)", 5, 1);
//                            formula       rows cols
```

Signature: `setArrayFormula(formula, rows, cols)`. The top-left cell
holds the formula; the result spills into `rows × cols`.

Equivalently:

```java
Worksheet ws = wb.getWorksheets().get(0);
ws.getCells().get("G2").setArrayFormula("=SUM(A2:D2*A2:D2)", 5, 1);
```

The top-left cell receives the array formula; results spill into
`rows × cols`. The cell is flagged as an array formula so the engine
re-evaluates after edits.

## Dynamic-array Excel 2021+ formulas

Aspose.Cells writes `=UNIQUE(...)` / `=FILTER(...)` as standard
formulas; Excel expands them in the UI. To force the legacy CSE
behaviour, wrap in `setArrayFormula`.

## R1C1 formula

```java
ws.getCells().get("A1").setR1C1Formula("=R[-1]C+R[-1]C[1]");
```

Mixed references use `[n]` (offset n columns forward) and `-n`
(backwards); absolute references use plain digits: `R1`, `C5`.

## Reading R1C1 from an Excel file

There is no direct getR1C1 accessor on the workbook — call
`Cell.getR1C1Formula()` (and `setR1C1Formula(String)`) on the cell
instance:

```java
String r1c1 = ws.getCells().get("A1").getR1C1Formula();
```

## Pitfalls

- **`setArrayFormula` mutates `rows × cols` cells** — the user loses
  data in that span.
- **Excel 2021 dynamic arrays** render *outside* the sheet's stored
  dimensions. Aspose does not validate spill conflicts; the saved
  file may disagree with what Excel displays.
- **R1C1 offsets are zero-based internally** — for `=R[-1]C`, `[-1]`
  means "row above the formula cell".
- **Don't mix A1 and R1C1** in the same save — Excel keeps both but
  the model is confusing.

## Related

- [set-get-formula.md](set-get-formula.md)
- [calculate-formulas.md](calculate-formulas.md)
