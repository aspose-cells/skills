## Verifying the docs

The generated skill includes a **verifier** that cross‑checks every API name, signature, and enum member against the real Aspose.Cells.GridJs assemblies. Running the verifier ensures that the documentation never drifts from the actual library, preventing hallucinations in AI‑generated code.

### How to run the verifier

```powershell
# From the root of the skill repository
python verifier/name_verifier.py .
```

The script will:

1. Parse all ````python` and ````typescript` code blocks in the markdown files.  
2. Reflect over the compiled **Aspose.Cells.GridJs** Python wheels and the **@aspose/cells-gridjs** npm package.  
3. Compare each discovered type, method, property, enum, and event against the documentation.  
4. Report any mismatches, missing members, or signature differences.

### Interpreting the results

- **0 errors** – The documentation is in sync with the library. You can safely ship the skill.  
- **>0 errors** – Edit the markdown files to correct the mismatched entries, then re‑run the verifier.  

### Continuous integration

Add the following step to your CI pipeline (e.g., GitHub Actions) to enforce zero‑error verification on every pull request:

```yaml
- name: Verify Aspose.Cells.GridJs docs
  run: |
    python -m pip install -r requirements.txt
    python verifier/name_verifier.py .
```

Running the verifier as part of CI guarantees that any future changes to the library (new methods, renamed parameters, etc.) are immediately reflected in the skill documentation, keeping AI agents reliable.
```markdown
## Verifying the docs

The **verifier** is a lightweight Python utility that parses the generated Markdown files, extracts every API name, signature, and enum member, and cross‑checks them against the compiled Aspose.Cells.GridJs Python package (`asposecellsgridjs`). This guardrail ensures that the skill documentation never drifts from the actual library implementation.

### How to run the verifier locally

```powershell
# From the root of the skill repository
python verifier/name_verifier.py .
```

The script will:

1. **Collect** all code fences (` ```python `, ` ```typescript `) from the Markdown files.
2. **Parse** each signature using a tolerant regex that respects optional parameters, default values, and overloads.
3. **Import** the `asposecellsgridjs` package (or the local wheel built from the source) and introspect its modules via `inspect`.
4. **Report** any mismatches:
   * Missing classes, methods, properties, or enums.
   * Parameter name or type discrepancies.
   * Signature differences (e.g., extra/removed optional arguments).

A successful run ends with:

```
✅ All documented APIs match the installed Aspose.Cells.GridJs package.
```

### CI integration

Add the verifier step to your CI pipeline (GitHub Actions, Azure Pipelines, etc.):

```yaml
name: Verify Skill Docs

on:
  push:
    paths:
      - '**/*.md'
  pull_request:
    paths:
      - '**/*.md'

jobs:
  verify-docs:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - name: Set up Python
        uses: actions/setup-python@v5
        with:
          python-version: '3.11'
      - name: Install package and verifier
        run: |
          python -m pip install --upgrade pip
          pip install -e .   # installs the Aspose.Cells.GridJs wheel in editable mode
          pip install -r verifier/requirements.txt
      - name: Run verifier
        run: python verifier/name_verifier.py .
```

If the verifier fails, the CI job will be marked **failed**, preventing the merge until the documentation is corrected.

### Troubleshooting common issues

| Symptom | Likely cause | Fix |
|---------|--------------|-----|
| `ImportError: No module named asposecellsgridjs` | The package isn’t installed in the CI environment. | Ensure the `pip install -e .` step points to the correct wheel or source directory. |
| `Signature mismatch for Spreadsheet.loadData` | The Markdown signature is out‑of‑date (e.g., a new optional parameter was added). | Update the method signature in **03-client-api-reference.md** to match the latest library version. |
| `Enum member not found: ExportFormat.PDFX` | The enum was renamed or removed in a newer release. | Verify the enum definition in the library source and adjust the table in **06-configuration.md** or **02-server-api-reference.md** accordingly. |

### Keeping the verifier up‑to‑date

When the Aspose.Cells.GridJs library releases a new version:

1. **Update the package** in the repository (replace the wheel or adjust the `requirements.txt` version).
2. **Run the verifier** locally to surface any documentation gaps.
3. **Edit the affected Markdown files** (add new methods, adjust parameter defaults, update enum tables).
4. **Commit** the changes; the CI verifier will confirm consistency.
