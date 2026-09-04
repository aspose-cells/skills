---
name: aspose-cells-java
description: Generate Java code that uses Aspose.Cells for Java to read, write, convert, and render Excel workbooks (XLSX, XLS, CSV, PDF, HTML, images). Use when the user asks to create, modify, save, convert, render, or otherwise manipulate Excel/spreadsheet files in Java using the Aspose.Cells library. Covers licenses, charts, pivot tables, formulas, formatting, rendering, encryption, and Linux/Docker font setup. Do NOT use for Aspose.Cells for .NET, Python, or any other language binding.
applies_to: Java

---

# Aspose.Cells for Java — Coding Skill

Authoritative reference for AI assistants generating Java code with
`com.aspose.cells.*`. Follow these rules and read the linked sub-skills
**before** emitting code.

## Non-negotiable rules (apply to every code block)

1. **Imports** — always `import com.aspose.cells.*;`. Never mix with POI
   (`org.apache.poi.*`) or any other Excel library in the same file.
2. **JavaBean accessors only** — `cell.getValue()`, `cell.setValue(...)`,
   `cell.putValue(...)`. Prefer `putValue` for typed literals
   (int/double/boolean/String/DateTime have dedicated overloads);
   `setValue(Object)` still works via auto-boxing. Never `cell.Value`,
   `cell.set_Value`, or other C#-style access.
3. **Java 8+ syntax** — no `var`, no records unless strictly required;
   target JDK 8 as the minimum supported version.
4. **All jars from the official Maven repo**, never manually copied.
   See [getting-started/maven-gradle-dependencies.md](getting-started/maven-gradle-dependencies.md).
5. **Apply a license** before any user-facing workbook construction
   (unless user is evaluating). See [_shared/license.md](_shared/license.md).
6. **PDF/image conversion on Linux/Docker** requires fonts + headless mode.
   See [_shared/fonts-linux-docker.md](_shared/fonts-linux-docker.md).
7. **Large files (>100k rows)** require `MemorySetting.MEMORY_PREFERENCE`
   or the `LightCells` API. See [_shared/large-files-memory.md](_shared/large-files-memory.md).
8. **After every `setFormula(...)`** call, run `workbook.calculateFormula()`
   (or `worksheet.calculateFormula(...)`) before reading the result or saving.
9. **Save with options object** — never `workbook.save("x.pdf")` for PDF,
   HTML, image, or CSV. Always pass a `*SaveOptions` instance.
10. **CSV/TSV exports only the active sheet by design.** Use
    `TxtSaveOptions.setExportAllSheets(true)` or iterate manually.

If any rule is violated by generated code, regenerate the snippet.

## How this skill is organised

```
skills/
├── SKILL.md                       ← this file (read first)
├── _shared/                       ← cross-cutting: license, fonts, memory, imports
├── getting-started/               ← install, open, create, save, settings
├── cells/                         ← cell values, formatting, sort, filter, validation
├── formulas/                      ← set/calc/array/R1C1 formulas, named ranges
├── worksheet/                     ← worksheets, page setup, panes, views
├── charts/                        ← all chart tasks
├── pivot-tables/                  ← pivot table lifecycle
├── rendering/                     ← PDF, image, HTML, CSV/JSON output
├── tables/                        ← ListObject tables
├── images-shapes/                 ← pictures, shapes, text boxes, OLE
├── security/                      ← encryption, digital signature, document properties
├── macros/                        ← VBA macros (.xlsm / .xlsb)
└── slicers-sparklines/            ← slicers and in-cell sparklines
```

Progressive loading — read `_shared/` plus the relevant category
folder only. Do **not** dump every sub-skill into context.

## Selecting a sub-skill

| Task                              | Read                                                       |
| --------------------------------- | ---------------------------------------------------------- |
| Install / Maven / Gradle          | `getting-started/maven-gradle-dependencies.md`             |
| Apply / verify a license          | `getting-started/setup-and-licensing.md`, `_shared/license.md` |
| New empty workbook                | `getting-started/create-workbook.md`                       |
| Open existing file                | `getting-started/open-file.md`                             |
| Save or convert format            | `getting-started/save-export.md`                           |
| Detect unknown file format        | `getting-started/detect-format.md`                         |
| Read or write cell value          | `cells/read-write-values.md`                               |
| Style / format cells              | `cells/cell-formatting-styles.md`                          |
| Conditional formatting            | `cells/conditional-formatting.md`                          |
| Insert / delete rows / columns    | `cells/rows-columns-operations.md`                         |
| Merge / unmerge                   | `cells/merge-unmerge-cells.md`                             |
| Sort                              | `cells/sort-data.md`                                       |
| Auto filter                       | `cells/auto-filter.md`                                     |
| Data validation (dropdowns)       | `cells/data-validation.md`                                 |
| Find / search / replace           | `cells/find-search-replace.md`                             |
| Smart markers / template bind     | `cells/smart-markers-reporting.md`                       |
| Formula set / get / calculate     | `formulas/set-get-formula.md`, `formulas/calculate-formulas.md` |
| Array or R1C1 formula             | `formulas/array-r1c1-formula.md`                           |
| Named range                       | `formulas/named-ranges.md`                                 |
| Add / rename / copy worksheets    | `worksheet/manage-worksheets.md`                           |
| Page setup, margins, print area   | `worksheet/page-setup-print.md`                            |
| Freeze panes, zoom, visibility    | `worksheet/panes-zoom-views.md`                            |
| Create any chart                  | `charts/create-chart.md`                                   |
| Pick a chart type                 | `charts/chart-types-selection.md`                          |
| Chart data / series formatting    | `charts/chart-data-formatting.md`                          |
| Render chart to image / PDF       | `charts/chart-to-image-pdf.md`                             |
| Create pivot table                | `pivot-tables/create-pivot-table.md`                       |
| Pivot field / value configuration | `pivot-tables/pivot-fields-formatting.md`                  |
| Refresh pivot cache / data        | `pivot-tables/refresh-pivot-table.md`                      |
| Convert workbook → PDF            | `rendering/convert-to-pdf.md`                              |
| Worksheet → PNG / JPEG / TIFF     | `rendering/worksheet-image.md`                             |
| Convert → HTML                    | `rendering/convert-html.md`                                |
| Export CSV / TSV / JSON / TXT     | `rendering/csv-json-txt-export.md`                         |
| ListObject (Excel Table)          | `tables/list-object-tables.md`                             |
| Insert images                     | `images-shapes/insert-image.md`                            |
| Shapes / textboxes                | `images-shapes/shapes-textboxes.md`                        |
| Embed OLE object (Word/PDF/…)     | `images-shapes/ole-objects.md`                             |
| Encrypt or decrypt file           | `security/encryption-decryption.md`                        |
| Digital signature                 | `security/digital-signature.md`                            |
| Built-in / custom doc properties  | `security/document-properties.md`                          |
| VBA macros (.xlsm / .xlsb)        | `macros/vba-macros.md`                                     |
| Pivot / table slicers             | `slicers-sparklines/slicers.md`                            |
| In-cell sparklines                | `slicers-sparklines/sparklines.md`                         |

When uncertain which sub-skill applies, ask the user — never invent APIs.

## Pitfalls to surface automatically

These are mistakes the research shows customers hit most. Inject the
correction proactively; do not wait for the user to discover them.

- **License missing** → generates watermark + 100-file limit.
- **PDF rendering on Linux** → squares ("tofu") or
  `RuntimeException: Fontconfig head is null` if fonts are absent.
- **`setFormula` returns blank value** → forgot `calculateFormula()`.
- **`workbook.save("file.pdf")` ignores options** → use `PdfSaveOptions`.
- **`refreshData()` on pivot** → deprecated; use
  `workbook.refreshAll()` / `worksheet.refreshPivotTables()` /
  `pivotTable.getPivotCache().refresh()` (since v25.x).
- **`cell.Value` / `cell.get_Value`** → use JavaBean `getValue()` /
  `setValue(...)` only.
- **CSV export missing sheets** → only active sheet is exported; set
  `TxtSaveOptions.setExportAllSheets(true)`.
- **Smart markers stay literal** → `&=...` still appears in cells after
  `process()`. Source name passed to `setDataSource(...)` must match
  the marker's source segment exactly, and the source must implement
  `ICellsDataTable` (the Java equivalent of .NET's `DataTable`).
- **Smart-marker totals show 0** → formulas below inserted rows were
  not recomputed; enable `designer.setCalculateFormula(true)` or call
  `workbook.calculateFormula()` before saving.
- **Chart image / PDF is blank** → render before `chart.calculate()`.
  For pivot-bound charts, `refreshPivotData()` first.
- **Pivot body is empty after source edits** → recompute with
  `pt.calculateData(new PivotTableCalculateOption().setRefreshData(true))`,
  not the obsolete `refreshData(); calculateData();` pair.
- **Circular formulas compute to 0** → iterative calculation is off
  by default; enable with
  `wb.getSettings().getFormulaSettings().setEnableIterativeCalculation(true)`,
  set `MaxIteration` and `MaxChange`, then `calculateFormula()`.
- **Text numbers stay text** → `putValue("123")` is a string; parse
  on write or call `cells.convertStringToNumericValue()` for a whole
  sheet.
- **Saved file is corrupt** → too many unique cell formats (>64,000);
  reuse styles via `applyColumnStyle` / `applyRowStyle` / `applyStyle`
  rather than allocating per cell.
- **License OK on dev, watermark on server** → subscription expired
  relative to the JAR's release date, or the `.lic` isn't on the
  classpath. Check `License.getSubscriptionExpireDate()` and load via
  `getClass().getResourceAsStream("/Aspose.Cells.Java.lic")`.

## Out of scope

- .NET, Python, Node.js, C++ bindings of Aspose.Cells — use the
  product-specific skill for those.
- Pure Microsoft Office interop (Jacob, POI-HSSF) — wrong library.
- Questions about pricing or licensing terms — point the user to
  <https://purchase.aspose.com>.
