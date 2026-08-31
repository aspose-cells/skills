# Aspose.Cells for .NET — SKILL repository

A maintainable, self-verifying SKILL for writing C#/.NET code that uses
[Aspose.Cells for .NET](https://reference.aspose.com/cells/net/): create, read,
edit, convert, or render spreadsheets (xlsx, xls, xlsm, csv, tsv, ods, and
spreadsheet-to-PDF/HTML/image).

Every snippet and inline API reference is verified by reflection against the
pinned `Aspose.Cells` 26.7.0 package, so a stale or misspelled API is caught by
the verifier, not by whatever user runs the skill.

## Repository layout

```
aspose-cells-net/                ← the repository root (the skill content)
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
└── verifier/                    ← guarantees the skill stays correct
    ├── Program.cs               ← reflection scan: checks every `inline` API
    │                               token, argument types, and variable types
    │                               against the pinned Aspose.Cells
    ├── CompileVerifier.cs       ← Roslyn-compiles every ```csharp block
    └── AsposeCellsVerifier.csproj ← the package version everything pins to
```

## How an agent uses this skill

This is a SKILL for coding agents (opencode, Claude Code, etc.), not a library
for humans. The agent does not "install" anything; it reads the docs and routes
to the right reference file.

1. **Load the skill.** The consuming agent registers this repository as a skill
   and reads `SKILL.md` first (its frontmatter `name`/`description` drive when
   the skill is triggered — e.g. whenever an `Aspose.Cells` task appears).
2. **Check version.** `SKILL.md` → `## Version assumptions` pins **Aspose.Cells
   26.7.0**. If the user's project uses a different version, the agent re-runs
   the verifiers (below) before trusting any snippet.
3. **Route to a reference.** When a task maps to one of the 17 golden rules or
   a row of the `When to read which reference` table, read that
   `references/*.md` file for the full pattern and pitfalls **before** writing
   code.
4. **Answer from the cheat sheet.** For quick lookups, use the `Cheat sheet`
   table — each row is one line of code plus the owning reference file.
5. **Never invent API.** If an API is not in the docs, the agent checks the
   pinned package via reflection or the official API reference, and if the docs
   were wrong, fixes the reference file and re-runs the verifiers.

## Verifying the docs

Two verifiers guard every C# snippet and inline reference in this skill. They
both run against the pinned Aspose.Cells package; exit code 0 means clean.

```powershell
# in the skill repo root
# 1) Compile every ```csharp block with Roslyn (catches wrong overloads,
#    bad argument types, misspelled members).
dotnet run --project verifier/AsposeCellsVerifier.csproj -c Release -- --compile

# 2) Reflection scan of inline references (checks that every type/member
#    named in prose and tables actually exists in 26.7, and that argument
#    types match the method signature).
dotnet run --project verifier/AsposeCellsVerifier.csproj -c Release
```

The compile run prints `COMPILE OK: N   ERRORS: M   BLOCKS: K`. Any `ERRORS`
line is a real doc bug - the report maps each error back to its source .md and
line. All blocks are compiled into ONE Roslyn compilation unit, so a type
declared in one block (e.g. `class Product`) is visible to later blocks that
use it. A `var wb`/`sheet`/`cells` referenced without a declaration is
auto-injected.

The reflection run prints `OK: n NOT FOUND: m SKIPPED: k CTOR-CHECKED: t`.
Exit code 0 means no NOT FOUND. A `NOT FOUND` line is a real doc bug -
either fix the doc or, if the API is genuinely gone, update the reference.
`SKIPPED` counts expressions the scanner cannot prove (loop variables,
`Console`, `typeof`, multi-line chains) - review them by eye, not by script.
The verifier resolves variable declarations (`Workbook wb = ...`,
`var sheet = ...`) across a code block and uses the resolved types for
argument-type checking, so `cells.ExportDataTable("0", 0, 10, 3, true)` is
correctly flagged as a type mismatch. When upgrading the package version,
bump it in `verifier/AsposeCellsVerifier.csproj` and re-run both verifiers.

## Adding or editing a reference

1. Pick a task/feature area; add or edit the matching file under `references/`.
2. Link the file from SKILL.md: add a **Golden rule**, a row in the **When to
   read which reference** table, and one or two **Cheat sheet** rows.
3. Re-run both verifiers. Both must exit 0 before the change is done.