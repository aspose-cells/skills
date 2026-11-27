# Aspose.Cells for Node.js via Java — SKILL repository

A self-verifying SKILL for writing CommonJS code with
[Aspose.Cells for Node.js via Java](https://docs.aspose.com/cells/nodejs-java/getting-started/)
(the `aspose.cells` npm package): create, read, edit, convert, or render spreadsheets
(xlsx, xls, xlsm, csv, tsv, ods, and spreadsheet-to-PDF/HTML/image).

This skill targets the **Java** binding. It is not Aspose.Cells for Node.js via C++
(`aspose.cells.node`), and not GridJs or GridWeb.

## Repository layout

```
aspose-cells-nodejs-via-java-skill/   ← the repository root (the skill content)
├── SKILL.md                     ← entry point — read first
├── references/                  ← one file per task / feature area
│   ├── cell-values-and-data.md
│   ├── charts.md
│   ├── conversion-and-rendering.md
│   ├── cross-platform-deployment.md
│   ├── digital-signatures-and-document-properties.md
│   ├── formulas-and-calculation.md
│   ├── images-and-shapes.md
│   ├── ole-objects.md
│   ├── performance-and-large-files.md
│   ├── pivot-tables.md
│   ├── setup-and-licensing.md
│   ├── slicers-and-sparklines.md
│   ├── smart-markers-reporting.md
│   ├── styles-and-formatting.md
│   ├── vba-macros.md
│   └── worksheets-rows-columns.md
└── verifier/
    └── verify.js                ← frontmatter, link, and JavaScript syntax check;
                                   optional runtime smoke test with --runtime
```

## How an agent uses this skill

This is a SKILL for coding agents, not a library for humans. The agent reads the
docs and routes to the right reference file.

1. **Load the skill.** The consuming agent registers this repository as a skill and
   reads `SKILL.md` first. Its frontmatter `name`/`description` drive when the skill
   is triggered — for example whenever an `aspose.cells` task appears.
2. **Check version.** `SKILL.md` → `## Version assumptions` pins the package
   version. If the user's project uses a different version, re-run the verifier
   before trusting version-sensitive overloads.
3. **Route to a reference.** When a task maps to a golden rule or a row of the
   reference routing table, read that `references/*.md` file before writing code.
4. **Respect the binding's limits.** No Java-array or `ResultSet` bulk imports,
   no printing, and no assumed JavaScript-to-Java conversion for dates, buffers,
   or Node streams.
5. **Never invent API.** If an API is not in the docs, check the installed package
   or the official Node.js API reference; if the docs were wrong, fix the reference
   file and re-run the verifier.

## Verifying the docs

```bash
# in the skill repo root
node verifier/verify.js
```

Static validation needs only Node.js. It checks frontmatter, local Markdown links,
and that every ```javascript block parses.

Runtime validation additionally needs `aspose.cells` and a working JDK:

```bash
node verifier/verify.js --runtime
```

It creates a workbook in the OS temp directory, writes values and a formula,
saves, reopens, and checks the calculated result. Set `ASPOSE_CELLS_LICENSE` to a
license path to run it licensed; otherwise the smoke test runs in evaluation mode.

## Adding or editing a reference

1. Pick a task/feature area; add or edit the matching file under `references/`.
2. Link it from `SKILL.md`: add a golden rule when it prevents a real bug, and a
   row in the reference routing table.
3. Re-run `node verifier/verify.js`. It must exit 0 before the change is done.

Do not add a file unless it removes real ambiguity for an AI assistant. Prefer
editing an existing reference to adding a new one.
