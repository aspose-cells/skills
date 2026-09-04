---
name: find-search-replace
description: Find, search, and replace cell text or formulas in Aspose.Cells for Java — exact matches, regex, whole-word, and case sensitivity. Use whenever the user wants to scan a sheet for a value or do search-and-replace.
applies_to: Java
keywords: [find, search, replace, regex, lookup]
---

# Find / search / replace

## Imports

```java
import com.aspose.cells.*;
```

## Quick example

```java
Worksheet ws = wb.getWorksheets().get(0);
FindOptions opts = new FindOptions();
opts.setCaseSensitive(false);
opts.setLookAtType(LookAtType.CONTAINS);

Cell cell = ws.getCells().find("Invoice", null, opts);

if (cell != null) {
    System.out.println("Found at " + cell.getName());   // e.g. C12
}
```

## Find next after the first hit

```java
FindOptions opts = new FindOptions();
opts.setCaseSensitive(false);
opts.setLookAtType(LookAtType.CONTAINS);
Cell first = ws.getCells().find("Invoice", null, opts);
Cell next = ws.getCells().find("Invoice", first, opts);
```

The previous hit's cell is passed as the starting cell to skip
matches before it.

## Replace all

```java
ReplaceOptions opts = new ReplaceOptions();
opts.setCaseSensitive(false);
opts.setMatchEntireCellContents(false);
int count = wb.replace("Invoice", "Receipt", opts);   // count of replacements
```

`replace` returns the number of replacements performed.

## Regex / formula lookups

```java
FindOptions opts = new FindOptions();
opts.setRegexKey(true);
opts.setLookAtType(LookAtType.ENTIRE_CONTENT);
Cell match = ws.getCells().find("[A-Z]{2}\\d{4}", null, opts);
```

To look **inside formulas** instead of values:

```java
FindOptions opts = new FindOptions();
opts.setLookInType(LookInType.ONLY_FORMULAS);   // search formula text, not values
Cell inFormula = ws.getCells().find("SUM(", null, opts);
```

## Pitfalls

- **`find` returns the first cell only.** Loop with `find(value, last, opts)`
  to enumerate all matches.
- **`LookInType.VALUES` skips formulas** that have a numeric result —
  use `LookInType.VALUES_EXCLUDE_FORMULA_CELL` for the opposite.
- **Regex lookbehind is not supported** by Aspose's NFA-compatible
  engine.
- **`replace` matches on the value of cells**; it does not rewrite
  formula strings. To swap references inside formulas, search with
  `opts.setLookInType(LookInType.ONLY_FORMULAS)` first.
- **Thread-safety** — `find`/`replace` mutate the worksheet; do not
  call them concurrently with other writes.

## Related

- [read-write-values.md](read-write-values.md)
