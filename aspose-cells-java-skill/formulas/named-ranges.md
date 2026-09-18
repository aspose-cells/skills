---
name: named-ranges
description: Create, list, edit, and remove workbook / worksheet scoped named ranges (Defined Names) with Aspose.Cells for Java. Use for VBA-style labels, named formula inputs, formula references, and integration-friendly aliases.
applies_to: Java
keywords: [named range, defined name, scope]
---

# Named ranges (defined names)

## Imports

```java
import com.aspose.cells.*;
```

## Quick example

```java
NameCollection names = wb.getWorksheets().getNames();
int idx = names.add("TaxRate");
names.get(idx).setRefersTo("=0.2");
names.get(idx).setRefersTo("=Sheet1!$B$1");           // after change
```

A defined name with a constant `=0.2` is a value-only name. Pointing
to a range or another formula makes it a referenceable range.

## Use in formulas

```java
ws.getCells().get("C1").setFormula("=A1*TaxRate");
```

The name resolves like any other cell reference.

## List / iterate

```java
NameCollection all = wb.getWorksheets().getNames();
for (int i = 0; i < all.getCount(); i++) {
    Name n = all.get(i);
    System.out.printf("%s = %s (sheet-index %d)%n",
        n.getFullText(), n.getRefersTo(), n.getSheetIndex());
}
```

## Worksheet-scoped name

```java
Worksheet ws = wb.getWorksheets().get(0);
NameCollection all = wb.getWorksheets().getNames();
int s = all.add("LocalTotal");
all.get(s).setSheetIndex(ws.getIndex());           // scope to the first sheet
all.get(s).setRefersTo("=SUM(A1:A10)");
```

Workbook-scoped names are referenced by simple `"TaxRate"`; sheet
scoped by `"Sheet1!LocalTotal"`. Use `Name.setSheetIndex(...)` to
mark a name as sheet-scoped.

## Pitfalls

- **`setRefersTo` accepts a formula with leading `=`** — values like
  `setRefersTo("0.2")` (without `=`) are treated as broken references.
- **Reserved names** (`Print_Area`, `Print_Titles`, `Database`, etc.)
  cannot be reused; choose another.
- **Rename breaks formulas** — Excel/Aspose don't rewrite formula
  text when a used range is renamed.
- **Names with spaces** — quote them in formulas:
  `"=\"My Range\""` or use brackets `[My Range]`.

## Related

- [set-get-formula.md](set-get-formula.md)
- [manage-worksheets.md](../worksheet/manage-worksheets.md)
