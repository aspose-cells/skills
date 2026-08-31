# Aspose.Cells for Python via .NET — SKILL repository

A maintainable, self-verifying SKILL for writing Python code that uses
[Aspose.Cells for Python via .NET](https://products.aspose.com/cells/python-net/)
(the `aspose-cells-python` package): create, read, edit, convert, or render spreadsheets
(xlsx, xls, xlsm, csv, tsv, ods, and spreadsheet-to-PDF/HTML/image).


Every snippet and inline API reference is verified against the installed
`aspose-cells-python` package, so a stale or misspelled API is caught by the verifier,
not by whatever user runs the skill.


## Repository layout


```
aspose-cells-python-skill/       ← the repository root (the skill content)
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
    ├── verify_syntax.py         ← AST-compile every ```python block
    ├── verify_methods.py        ← stub-based method/attribute existence check
    └── verify_runtime.py        ← execute every block against real fixtures
```


## How an agent uses this skill


This is a SKILL for coding agents (opencode, Claude Code, etc.), not a library
for humans. The agent does not "install" anything; it reads the docs and routes
to the right reference file.


1. **Load the skill.** The consuming agent registers this repository as a skill
   and reads `SKILL.md` first (its frontmatter `name`/`description` drive when
   the skill is triggered — e.g. whenever an `aspose.cells` task appears).
2. **Check version.** `SKILL.md` → `## Version assumptions` pins the
   `aspose-cells-python` package version. If the user's project uses a different version,
   the agent re-runs the verifiers (below) before trusting any snippet.
3. **Route to a reference.** When a task maps to one of the golden rules or a row
   of the `When to read which reference` table, read that `references/*.md` file
   for the full pattern and pitfalls **before** writing code.
4. **Answer from the cheat sheet.** For quick lookups, use the `Cheat sheet`
   table — each row is one line of code plus the owning reference file.
5. **Never invent API.** If an API is not in the docs, the agent checks the
   installed package via reflection or the official API reference, and if the
   docs were wrong, fixes the reference file and re-runs the verifiers.


## Verifying the docs


Three verifiers guard every Python snippet and inline reference in this skill.
They run against the installed `aspose-cells-python` package; exit code 0 means clean.


```bash
# in the skill repo root

# 1) Syntax check: AST-compile every ```python block (catches indentation
#    errors, unbalanced braces, bad string escapes — no runtime needed).
python verifier/verify_syntax.py

# 2) Method/attribute check: parse the shipped .pyi type stubs, resolve
#    receiver types in each code block via AST, and flag calls to members
#    that don't exist on that type (falls back to runtime hasattr).
python verifier/verify_methods.py

# 3) Runtime check: execute every ```python block against auto-generated
#    fixture files (xlsx, png, json, pfx, lic…) and report any exception.
#    This is the only layer that catches wrong argument types, overload
#    disambiguation failures, and subclassability issues.
python verifier/verify_runtime.py
```


**Layer summary**

| Verifier | What it catches | Speed | Needs runtime? |
|----------|----------------|-------|----------------|
| `verify_syntax.py` | Syntax errors (bad indentation, unclosed parens) | instant | No |
| `verify_methods.py` | Hallucinated method/attribute names (e.g. `r1_c1_formula`) | ~1 s | Yes (import only) |
| `verify_runtime.py` | Wrong arg types, overload mismatches, subclassing failures | ~30 s | Yes (full exec) |


The runtime run prints `blocks executed: N   blocks with exceptions: M`. Any
`ERR` line is either a real doc bug or an environmental issue (missing fixture,
platform-specific cert loading, placeholder license key). Real doc bugs show
`TypeError` / `AttributeError` about a call; environmental failures show
`FileNotFoundError`, `CellsException: The file is corrupted`, or
`InvalidOperationException: Authentication failed`.


## Adding or editing a reference


1. Pick a task/feature area; add or edit the matching file under `references/`.
2. Link the file from SKILL.md: add a **Golden rule**, a row in the **When to
   read which reference** table, and one or two **Cheat sheet** rows.
3. Re-run all three verifiers. All must exit 0 before the change is done.
