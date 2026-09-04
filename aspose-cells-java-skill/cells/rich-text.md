---
name: rich-text
description: Build rich-text cells with multiple FontSettings runs in Aspose.Cells for Java — mixed bold/size/colour inside a single cell. Use when the user wants a single cell to contain differently-styled fragments.
applies_to: Java
keywords: [rich text, font run, formatting]
---

# Rich text

## Imports

```java
import com.aspose.cells.*;
```

## Quick example

```java
Cell c = ws.getCells().get("A1");
c.setValue("Plain bold red");

FontSetting run = new FontSetting(6, 8, wb.getWorksheets());  // start, length, worksheets
run.getFont().setBold(true);
run.getFont().setColor(com.aspose.cells.Color.getRed());

FontSetting[] runs = new FontSetting[]{run};
c.setCharacters(runs);
```

## Standalone rich-text cell

```java
Cell a2 = ws.getCells().get("A2");
a2.setValue("Total $1,234.56");

FontSetting head = new FontSetting(0, 6, wb.getWorksheets());  // "Total "
FontSetting tail = new FontSetting(6, 9, wb.getWorksheets());  // "$1,234.56"
tail.getFont().setBold(true);
tail.getFont().setColor(com.aspose.cells.Color.getGreen());

a2.setCharacters(new FontSetting[]{head, tail});
```

## Pitfalls

- **`setValue(...)` wipes previous rich-text runs** — call
  `setCharacters` after setting the base value, or use `setValue("")`
  to fully reset.
- **`FontSetting` constructor takes `(startIndex, length, WorksheetCollection)`** —
  text positions are byte offsets into the cell value, not separate
  strings.
- **HTML is not supported** inside rich-text runs. Convert HTML
  fragments to multiple `FontSetting` manually.
- **Each FontSetting** is fully independent — there is no cascade
  from cell style to runs.
- **Excel limits runs per cell** to 64 visible text segments.

## Related

- [cell-formatting-styles.md](cell-formatting-styles.md)
- [hyperlinks-comments.md](hyperlinks-comments.md)
