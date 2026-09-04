---
name: conditional-formatting
description: Add conditional formatting rules (cell value, formula, colour scale, data bar, icon set) to ranges in Aspose.Cells for Java. Use when the user asks for heat-map-style colouring, validation-driven highlights, or Excel-style icon sets.
applies_to: Java
keywords: [conditional formatting, color scale, data bar, icon set, highlight]
---

# Conditional formatting

## Imports

```java
import com.aspose.cells.*;
```

## Quick example — "between" highlight

```java
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().get(0);
Cells cells = ws.getCells();                   // unused; satisfies the wrapper
ConditionalFormattingCollection fcc = ws.getConditionalFormattings();
int idx = fcc.add();
FormatConditionCollection conditions = fcc.get(idx);
FormatCondition cond = conditions.get(0);

CellArea area = new CellArea();
area.StartRow = 0; area.EndRow = 99;
area.StartColumn = 0; area.EndColumn = 0;
conditions.addArea(area);
cond.setFormula1("100");
cond.setFormula2("500");
cond.setType(FormatConditionType.CELL_VALUE);
cond.setOperator(OperatorType.BETWEEN);
cond.getStyle().getFont().setColor(com.aspose.cells.Color.getRed());
```

## Colour scale

```java
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().get(0);
Cells cells = ws.getCells();                   // unused; satisfies the wrapper
ConditionalFormattingCollection fcc = ws.getConditionalFormattings();
FormatConditionCollection conditions = fcc.get(fcc.add());
FormatCondition cond = conditions.get(0);
cond.setType(FormatConditionType.COLOR_SCALE);
ColorScale cs = cond.getColorScale();
cs.setMinColor(com.aspose.cells.Color.getGreen());
cs.setMidColor(com.aspose.cells.Color.getYellow());
cs.setMaxColor(com.aspose.cells.Color.getRed());
```

## Data bar / Icon set

```java
Workbook wb = new Workbook();
Worksheet ws = wb.getWorksheets().get(0);
Cells cells = ws.getCells();                   // unused; satisfies the wrapper
ConditionalFormattingCollection fcc = ws.getConditionalFormattings();
FormatConditionCollection conditions = fcc.get(fcc.add());
FormatCondition cond = conditions.get(0);
cond.setType(FormatConditionType.DATA_BAR);
DataBar bar = cond.getDataBar();
bar.getMinCfvo().setType(FormatConditionValueType.MIN);
bar.getMaxCfvo().setType(FormatConditionValueType.MAX);
bar.setColor(com.aspose.cells.Color.getBlue());
```

## Pitfalls

- **Areas live on the `FormatConditionCollection`**, not on the
  `FormatCondition` itself — call `conditions.addArea(area)`.
- **`setFormula1`/`setFormula2` strings** must include leading `=`
  *only* when the operator is formula-driven (e.g. `EXPRESSION`); for
  cell-value operators do **not** prefix `=`.
- **`FormatConditionCollection` is per worksheet**, not global.
- **Styles on condition** — modify a brand-new `Style` derived from
  the cell, never reuse the existing style.
- **Rule deletion** — `fcc.removeAt(idx)` shifts other indexes.

## Related

- [cell-formatting-styles.md](cell-formatting-styles.md)
- [data-validation.md](data-validation.md)
