# Aspose.Cells for Java — AI Coding Skill

A structured knowledge base that lets any LLM (Claude, GPT, Gemini,
Copilot, Cursor) generate correct, current, production-ready Java code
for **Aspose.Cells for Java**.

## Why this exists

LLMs frequently produce broken Aspose.Cells code because:

- API names drift between releases.
- Java vs .NET surface area is similar but not identical (property
  getters, package names, overloads).
- Many tutorials are out of date.
- Edge-case pitfalls (Linux fonts, OOM, deprecated refresh APIs) are
  rarely in the training corpus.

This skill captures what the model needs and nothing it does not. It
follows the [Agent Skills spec](https://agentskills.io) — frontmatter
metadata per file, `SKILL.md` as the entry point, and supporting files
loaded only when relevant.

## Layout

```
skills/
├── SKILL.md                       read first, gives the rules + an index
├── _shared/                       license / fonts / memory / imports
├── getting-started/               install, open, save, settings
├── cells/                         cell values, formatting, sort, filter
├── formulas/                      formulas + calculation
├── worksheet/                     worksheets, page setup, panes
├── charts/                        all chart tasks
├── pivot-tables/                  pivot tables and refresh
├── rendering/                     PDF, image, HTML, CSV/JSON/TXT
├── tables/                        ListObject tables
├── images-shapes/                 pictures, shapes, text boxes
└── security/                      encryption, digital signature
```

Each sub-skill is a single Markdown file under ~150 lines, scoped to one
task. Load only the one(s) relevant to the current request — never the
entire tree.

## Conventions

- Language: Java 8 and above.
- Imports: `com.aspose.cells.*` only.
- All examples compile against current Aspose.Cells for Java.
- Examples assume a license is loaded (see `_shared/license.md`).
- For Docker / Linux rendering, see `_shared/fonts-linux-docker.md`.

## Contributing

1. Identify the customer task the new skill must serve.
2. Write the front-matter: `name`, `description`, `applies_to`, `since`.
3. Body order: **When to use → Imports → Quick example → Key APIs →
   Pitfalls → Related skills**.
4. Keep the example minimal and runnable.
5. Update the index table in `SKILL.md`.

Do not add a file unless it removes real ambiguity for an AI assistant.
Prefer editing an existing skill to adding a new one.
