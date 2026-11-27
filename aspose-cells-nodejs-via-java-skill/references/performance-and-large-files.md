# Performance and large files

## Load with a memory preference

```javascript
const cells = require("aspose.cells");

const options = new cells.LoadOptions();
options.setMemorySetting(cells.MemorySetting.MEMORY_PREFERENCE);
const workbook = new cells.Workbook("large.xlsx", options);
```

Memory preference trades some speed for lower memory use; it does not make an arbitrarily large workbook free to process. Set it when loading, before cells are materialized. Benchmark the real workload and watch the Node process plus JVM/native memory, not only the JavaScript heap.

## Reduce object-model work

- Batch value and formula edits, then calculate once.
- Reuse styles instead of creating one per cell.
- Avoid scanning empty rectangular areas with `cells.get(row, column)`.
- Prefer materialized-row iteration for reads and typed loops for writes.
- Avoid repeatedly auto-fitting large ranges after every edit.
- Render only requested sheets/pages when the output API supports selection.

The Java LightCells API can reduce memory for suitable sequential workloads, but it requires a Java interface implementation. Do not invent a plain JavaScript object with similarly named callbacks. Use LightCells only after confirming that the installed Node-Java bridge supports the required interface proxy and testing it against the exact release.

## Concurrency and isolation

A `Workbook` and its dependent objects are mutable and are not thread-safe. Give each job its own workbook and never share sheets, cells, styles, or renderers between concurrent requests.

For a Node service, isolate heavy jobs in separate worker processes or carefully bounded workers with their own module state. A worker limit prevents CPU and native-memory spikes; a queue prevents unbounded work from accumulating. Keep in mind that worker isolation does not make two operations on the same workbook safe.

## Reliability controls

Set request/file-size limits, operation timeouts, and process memory limits for untrusted files. Use job-specific output paths and clean up temporary files in `finally` logic. If the JVM or native bridge crashes, restart the whole worker process rather than trying to reuse partially initialized objects.

Measure these stages separately when diagnosing a timeout: file load, cell population, formula calculation, pivot/chart refresh, and rendering. The slow stage determines the fix; adding more memory does not repair an accidental whole-sheet scan.
