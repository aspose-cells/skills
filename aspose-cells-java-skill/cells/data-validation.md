---
name: data-validation
description: Add Excel-style data validation (dropdowns, numbers, dates, custom formulas) to cells in Aspose.Cells for Java. Use whenever the user wants to restrict input, show input messages, or validation alerts.
applies_to: Java
keywords: [validation, dropdown, list, input message, error alert]
---

# Data validation

## Imports

```java
import com.aspose.cells.*;
```

## Quick example — dropdown list

```java
Worksheet ws = wb.getWorksheets().get(0);
ValidationCollection validations = ws.getValidations();
CellArea area = new CellArea();
area.StartRow = 1; area.EndRow = 100;
area.StartColumn = 4; area.EndColumn = 4;
int idx = validations.add(area);                   // attach the range up-front (column E, rows 2–101)
Validation v = validations.get(idx);

v.setType(ValidationType.LIST);
v.setFormula1("Pass,Fail,Pending");
v.setShowInput(true);
v.setInputTitle("Status");
v.setInputMessage("Choose one of Pass, Fail, or Pending");
v.setShowError(true);
v.setErrorTitle("Invalid status");
v.setErrorMessage("Please use the dropdown.");
```

## Custom formula

```java
CellArea area = new CellArea();
area.StartRow = 1; area.EndRow = 100;
Validation v = ws.getValidations().get(ws.getValidations().add(area));
v.setType(ValidationType.CUSTOM);
v.setFormula1("=AND(ISNUMBER(B2), B2>0)");
```

## Read validations

```java
for (int i = 0; i < ws.getValidations().getCount(); i++) {
    Validation current = ws.getValidations().get(i);
    // iterate `current.getAreas()`
}
```

## Pitfalls

- **`addArea` does not erase previous areas** — multiple calls stack
  ranges onto one rule.
- **Dropdown list source** — `setFormula1("Pass,Fail")` works;
  passing `"=Sheet2!$A$1:$A$5"` switches to a range-based dropdown.
- **`ValidationType.LIST` is also `inCellDropdown`** — set
  `v.setInCellDropDown(true)` for the small arrow (default is true).
- **`Custom` validation** — formula must start with `=` and produce
  TRUE/FALSE.
- **Empty input** — set `v.setIgnoreBlank(true)` to allow blank cells
  past validation.

## Related

- [auto-filter.md](auto-filter.md)
- [conditional-formatting.md](conditional-formatting.md)
